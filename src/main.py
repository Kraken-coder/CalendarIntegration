
import os
import json
import httpx
import urllib.parse
from fastapi import FastAPI, Request, HTTPException, Depends, BackgroundTasks
from fastapi.responses import RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from supabase import create_client, Client
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime, timezone

load_dotenv()

app = FastAPI(title="Calendar Service V2 Integration")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SUPABASE_URL = os.getenv("SUPABASE_URL", "https://api.spiked.ai")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
supabase: Optional[Client] = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        print(f"Failed to initialize Supabase client: {e}")

DEFAULT_TEST_USER_ID = "fd3ff615-b248-4e8f-84f1-ff458bf30d48"

RECALL_API_KEY = os.getenv("RECALL_API_KEY")
PUBLIC_URL = os.getenv("PUBLIC_URL", "http://localhost:8000")
FRONTEND_URL = os.getenv("FRONTEND_URL", "http://localhost:5173")
RECALL_TRANSCRIPT_WEBHOOK_URL = os.getenv("RECALL_TRANSCRIPT_WEBHOOK_URL", f"{PUBLIC_URL}/webhook/recall/transcript")

GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CALENDAR_OAUTH_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CALENDAR_OAUTH_CLIENT_SECRET")
MS_CLIENT_ID = os.getenv("MICROSOFT_OUTLOOK_OAUTH_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MICROSOFT_OUTLOOK_OAUTH_CLIENT_SECRET")

RECALL_BASE_URL = "https://us-west-2.recall.ai"

def get_recall_client():
    return httpx.AsyncClient(base_url=RECALL_BASE_URL, headers={"Authorization": f"Token {RECALL_API_KEY}"})

def get_user_id_from_token(auth_header: str) -> Optional[str]:
    if not auth_header or not auth_header.startswith("Bearer "):
        return None
    token = auth_header.split(" ")[1]
    if supabase:
        try:
            res = supabase.auth.get_user(token)
            if res and res.user:
                return res.user.id
        except Exception:
            pass
    return None

@app.get("/integrations/calendar/google/initiate")
async def initiate_google_calendar(request: Request):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID # Fallback for demo
            
    state = json.dumps({"userId": user_id})
    scopes = ["https://www.googleapis.com/auth/userinfo.email", "https://www.googleapis.com/auth/calendar.events.readonly"]
    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "redirect_uri": f"{PUBLIC_URL}/oauth-callback/google-calendar",
        "response_type": "code",
        "scope": " ".join(scopes),
        "access_type": "offline",
        "prompt": "consent",
        "state": state
    }
    url = f"https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params)
    return {"auth_url": url}

@app.get("/oauth-callback/google-calendar")
async def google_calendar_callback(request: Request, code: str, state: str):
    state_data = json.loads(state)
    user_id = state_data.get("userId")
    
    async with httpx.AsyncClient() as client:
        token_res = await client.post("https://oauth2.googleapis.com/token", data={
            "client_id": GOOGLE_CLIENT_ID,
            "client_secret": GOOGLE_CLIENT_SECRET,
            "redirect_uri": f"{PUBLIC_URL}/oauth-callback/google-calendar",
            "grant_type": "authorization_code",
            "code": code
        })
        token_data = token_res.json()
        
        if "error" in token_data:
            return RedirectResponse(f"{FRONTEND_URL}/integrations?calendar_error={token_data['error']}")
            
        refresh_token = token_data.get("refresh_token")
        
    async with get_recall_client() as recall:
        recall_res = await recall.post("/api/v2/calendars/", json={
            "platform": "google_calendar",
            "webhook_url": f"{PUBLIC_URL}/webhooks/recall-calendar-updates",
            "oauth_refresh_token": refresh_token,
            "oauth_client_id": GOOGLE_CLIENT_ID,
            "oauth_client_secret": GOOGLE_CLIENT_SECRET
        })
        recall_calendar = recall_res.json()
        
    print(f"Recall calendar creation response: {recall_calendar}")
    
    if not supabase:
        print("ERROR: Supabase client is not initialized. Check your SUPABASE_URL and SUPABASE_KEY in .env.")
    elif not recall_calendar.get("id"):
        print(f"ERROR: No 'id' returned from Recall API. Full response: {recall_calendar}")
    else:
        print(f"Attempting to insert calendar for user {user_id} into Supabase...")
        try:
            response = supabase.table("calendars").insert({
                "platform": "google_calendar",
                "recall_id": recall_calendar.get("id"),
                "recall_data": recall_calendar,
                "user_id": user_id
            }).execute()
            print(f"Successfully saved to Supabase: {response.data}")
        except Exception as e:
            print(f"CRITICAL Error saving to Supabase DB: {e}")
            print(f"Exception type: {type(e)}")
        
    return RedirectResponse(f"{FRONTEND_URL}/integrations?calendar_success=true")

@app.get("/integrations/calendar/outlook/initiate")
async def initiate_outlook_calendar(request: Request):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
            
    state = json.dumps({"userId": user_id})
    scopes = ["offline_access", "https://graph.microsoft.com/Calendars.Read", "openid", "email"]
    params = {
        "client_id": MS_CLIENT_ID,
        "redirect_uri": f"{PUBLIC_URL}/oauth-callback/microsoft-outlook",
        "response_type": "code",
        "scope": " ".join(scopes),
        "state": state
    }
    url = f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params)
    return {"auth_url": url}

@app.get("/oauth-callback/microsoft-outlook")
async def outlook_calendar_callback(request: Request, code: str, state: str):
    state_data = json.loads(state)
    user_id = state_data.get("userId")
    
    async with httpx.AsyncClient() as client:
        token_res = await client.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", data={
            "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "redirect_uri": f"{PUBLIC_URL}/oauth-callback/microsoft-outlook",
            "grant_type": "authorization_code",
            "code": code
        })
        token_data = token_res.json()
        
        if "error" in token_data:
            return RedirectResponse(f"{FRONTEND_URL}/integrations?calendar_error={token_data['error']}")
            
        refresh_token = token_data.get("refresh_token")
        
    async with get_recall_client() as recall:
        recall_res = await recall.post("/api/v2/calendars/", json={
            "platform": "microsoft_outlook",
            "webhook_url": f"{PUBLIC_URL}/webhooks/recall-calendar-updates",
            "oauth_refresh_token": refresh_token,
            "oauth_client_id": MS_CLIENT_ID,
            "oauth_client_secret": MS_CLIENT_SECRET
        })
        recall_calendar = recall_res.json()
        
    print(f"Recall calendar creation response (Outlook): {recall_calendar}")
    
    if not supabase:
        print("ERROR: Supabase client is not initialized. Check your SUPABASE_URL and SUPABASE_KEY in .env.")
    elif not recall_calendar.get("id"):
        print(f"ERROR: No 'id' returned from Recall API. Full response: {recall_calendar}")
    else:
        print(f"Attempting to insert calendar for user {user_id} into Supabase...")
        try:
            response = supabase.table("calendars").insert({
                "platform": "microsoft_outlook",
                "recall_id": recall_calendar.get("id"),
                "recall_data": recall_calendar,
                "user_id": user_id
            }).execute()
            print(f"Successfully saved to Supabase (Outlook): {response.data}")
        except Exception as e:
            print(f"CRITICAL Error saving to Supabase DB (Outlook): {e}")
            print(f"Exception type: {type(e)}")
        
    return RedirectResponse(f"{FRONTEND_URL}/integrations?calendar_success=true")

# ===== HELPERS FOR CALENDAR EVENT SYNC & RULES =====

def get_event_title(event: dict) -> str:
    raw = event.get("raw") or {}
    platform = event.get("platform") or "google_calendar"
    if platform == "google_calendar":
        return raw.get("summary") or "Untitled Meeting"
    elif platform == "microsoft_outlook":
        return raw.get("subject") or "Untitled Meeting"
    return "Untitled Meeting"

async def fetch_recall_calendar_events(recall_id: str, last_updated_ts: Optional[str] = None) -> List[dict]:
    events = []
    params = {"calendar_id": recall_id}
    if last_updated_ts:
        params["updated_at__gte"] = last_updated_ts
        
    url = "/api/v2/calendar-events/"
    async with get_recall_client() as client:
        while url:
            if url.startswith("http"):
                if url.startswith("http:"):
                    url = url.replace("http:", "https:", 1)
                resp = await client.get(url)
            else:
                resp = await client.get(url, params=params)
                params = None
                
            resp.raise_for_status()
            data = resp.json()
            events.extend(data.get("results", []))
            url = data.get("next")
    return events

def get_calendar_email(calendar: dict) -> str:
    recall_data = calendar.get("recall_data") or {}
    email = recall_data.get("oauth_email") or recall_data.get("platform_email") or ""
    return email.lower()

def get_attendees_for_calendar_event(event: dict) -> List[dict]:
    attendees = []
    platform = event.get("platform", "google_calendar")
    raw = event.get("raw") or {}
    
    if platform == "google_calendar":
        raw_attendees = raw.get("attendees") or []
        for attendee in raw_attendees:
            email = attendee.get("email", "").lower()
            if email:
                attendees.append({
                    "email": email,
                    "accepted": attendee.get("responseStatus") == "accepted"
                })
        organizer = raw.get("organizer", {})
        org_email = organizer.get("email", "").lower()
        if org_email:
            attendees.append({
                "email": org_email,
                "accepted": True
            })
    elif platform == "microsoft_outlook":
        raw_attendees = raw.get("attendees") or []
        for attendee in raw_attendees:
            email_address = attendee.get("emailAddress", {})
            email = email_address.get("address", "").lower()
            status = attendee.get("status", {})
            response = status.get("response", "")
            if email:
                attendees.append({
                    "email": email,
                    "accepted": response in ["accepted", "organizer"]
                })
        organizer = raw.get("organizer", {})
        org_email_address = organizer.get("emailAddress", {})
        org_email = org_email_address.get("address", "").lower()
        if org_email:
            attendees.append({
                "email": org_email,
                "accepted": True
            })
    return attendees

def is_external_event(event: dict, calendar_email: str) -> bool:
    if not calendar_email:
        return False
    try:
        cal_domain = calendar_email.split("@")[1].lower()
    except IndexError:
        return False
        
    attendees = get_attendees_for_calendar_event(event)
    for attendee in attendees:
        email = attendee["email"]
        try:
            domain = email.split("@")[1].lower()
            if domain != cal_domain:
                return True
        except IndexError:
            continue
    return False

def is_confirmed_event(event: dict, calendar_email: str) -> bool:
    if not calendar_email:
        return False
    attendees = get_attendees_for_calendar_event(event)
    for attendee in attendees:
        if attendee["email"] == calendar_email.lower() and attendee["accepted"]:
            return True
    return False

def should_record_event(event: dict) -> bool:
    recall_data = event.get("recall_data") or {}
    settings = recall_data.get("settings") or {}
    should_record_automatic = settings.get("should_record_automatic", False)
    should_record_manual = settings.get("should_record_manual")
    
    if should_record_manual is not None:
        record_decision = should_record_manual
    else:
        record_decision = should_record_automatic
        
    meeting_url = recall_data.get("meeting_url")
    return bool(record_decision and meeting_url)

async def update_auto_record_status_for_events(calendar: dict, events: List[dict]):
    calendar_email = get_calendar_email(calendar)
    
    calendar_settings = (calendar.get("recall_data") or {}).get("settings") or {}
    auto_record_external = calendar_settings.get("auto_record_external_events", False)
    auto_record_only_confirmed = calendar_settings.get("auto_record_only_confirmed_events", False)
    
    for event in events:
        recall_data = event.get("recall_data") or {}
        
        # 1. Skip past events
        end_time_str = recall_data.get("end_time")
        if end_time_str:
            try:
                clean_end_time = end_time_str.replace("Z", "+00:00")
                end_time = datetime.fromisoformat(clean_end_time)
                if end_time < datetime.now(timezone.utc):
                    print(f"INFO: Ignoring event {recall_data.get('title')} as it has ended")
                    continue
            except Exception as e:
                print(f"Error checking end time: {e}")
                
        # 2. Evaluate auto-recording
        should_record_automatic = False
        if auto_record_external:
            should_record_automatic = is_external_event(recall_data, calendar_email)
            
        if auto_record_only_confirmed:
            should_record_automatic = should_record_automatic and is_confirmed_event(recall_data, calendar_email)
            
        # 3. Update settings on the event
        settings = recall_data.get("settings") or {}
        settings["should_record_automatic"] = should_record_automatic
        recall_data["settings"] = settings
        
        # Save to database
        supabase.table("calendar_events").update({"recall_data": recall_data}).eq("id", event["id"]).execute()
        print(f"INFO: Updated auto record status of '{recall_data.get('title')}' to {should_record_automatic}")

async def sync_event_bot_schedule(user_id: str, event: dict):
    recall_data = event.get("recall_data") or {}
    event_recall_id = event.get("recall_id")
    meeting_url = recall_data.get("meeting_url")
    
    record = should_record_event(event)
    
    async with get_recall_client() as client:
        if record:
            print(f"INFO: Schedule bot for event {event.get('id')}")
            start_time = recall_data.get("start_time", "")
            deduplication_key = f"{start_time}-{meeting_url}"
            
            bot_config = {
                "bot_name": "SpikedAI",
                "recording_config": {
                    "transcript": {"provider": {"meeting_captions": {}}},
                    "realtime_endpoints": [{
                        "type": "webhook",
                        "url": RECALL_TRANSCRIPT_WEBHOOK_URL,
                        "events": ["transcript.data", "transcript.partial_data"]
                    }]
                },
                "metadata": {
                    "user_id": user_id,
                    "calendar_event_id": event_recall_id
                }
            }
            
            try:
                resp = await client.post(
                    f"/api/v2/calendar-events/{event_recall_id}/bot/",
                    json={
                        "deduplication_key": deduplication_key,
                        "bot_config": bot_config
                    }
                )
                if resp.status_code in [200, 201]:
                    updated_event = resp.json()
                    settings = (event.get("recall_data") or {}).get("settings") or {}
                    updated_event["settings"] = settings
                    supabase.table("calendar_events").update({"recall_data": updated_event}).eq("id", event["id"]).execute()
                    print(f"Successfully scheduled bot for event {event_recall_id}")
                else:
                    print(f"Error scheduling bot for event {event_recall_id}: {resp.status_code} - {resp.text}")
            except Exception as e:
                print(f"Failed to schedule bot for event {event_recall_id}: {e}")
        else:
            print(f"INFO: Unschedule/Delete bot for event {event.get('id')}")
            try:
                resp = await client.delete(f"/api/v2/calendar-events/{event_recall_id}/bot/")
                if resp.status_code in [200, 204]:
                    fresh_resp = await client.get(f"/api/v2/calendar-events/{event_recall_id}/")
                    if fresh_resp.status_code == 200:
                        updated_event = fresh_resp.json()
                        settings = (event.get("recall_data") or {}).get("settings") or {}
                        updated_event["settings"] = settings
                        supabase.table("calendar_events").update({"recall_data": updated_event}).eq("id", event["id"]).execute()
                        print(f"Successfully unscheduled bot for event {event_recall_id}")
                else:
                    print(f"Delete bot status for event {event_recall_id}: {resp.status_code}")
            except Exception as e:
                print(f"Failed to delete bot for event {event_recall_id}: {e}")

async def delete_bot_for_event(event_recall_id: str):
    print(f"INFO: Delete bot for deleted recall event {event_recall_id}")
    async with get_recall_client() as client:
        try:
            await client.delete(f"/api/v2/calendar-events/{event_recall_id}/bot/")
        except Exception as e:
            print(f"Failed to delete bot for deleted event {event_recall_id}: {e}")


# ===== BACKGROUND WORKERS FOR CALENDAR PROCESSING =====

async def process_calendar_sync_background(recall_id: str, last_updated_ts: Optional[str] = None):
    calendar_res = supabase.table("calendars").select("*").eq("recall_id", recall_id).execute()
    if not calendar_res.data:
        print(f"Could not find calendar with recall_id: {recall_id}")
        return
    calendar = calendar_res.data[0]
    calendar_id = calendar["id"]
    user_id = calendar["user_id"]
    
    try:
        events = await fetch_recall_calendar_events(recall_id, last_updated_ts)
        print(f"Fetched {len(events)} events for calendar {recall_id}")
    except Exception as e:
        print(f"Failed to fetch calendar events from Recall: {e}")
        return
        
    for event in events:
        event_recall_id = event["id"]
        if event.get("is_deleted", False):
            supabase.table("calendar_events").delete().eq("recall_id", event_recall_id).execute()
            await delete_bot_for_event(event_recall_id)
        else:
            db_event_res = supabase.table("calendar_events").select("*").eq("recall_id", event_recall_id).execute()
            existing_event = db_event_res.data[0] if db_event_res.data else None
            
            settings = {"should_record_automatic": False, "should_record_manual": None}
            if existing_event:
                existing_recall_data = existing_event.get("recall_data") or {}
                existing_settings = existing_recall_data.get("settings") or {}
                settings["should_record_automatic"] = existing_settings.get("should_record_automatic", False)
                settings["should_record_manual"] = existing_settings.get("should_record_manual", None)
                
            event["title"] = get_event_title(event)
            event["settings"] = settings
            
            event_data = {
                "platform": event.get("platform", "google_calendar"),
                "recall_id": event_recall_id,
                "recall_data": event,
                "calendar_id": calendar_id
            }
            
            if existing_event:
                supabase.table("calendar_events").update(event_data).eq("id", existing_event["id"]).execute()
                event_db = {**existing_event, "recall_data": event}
            else:
                insert_res = supabase.table("calendar_events").insert(event_data).execute()
                event_db = insert_res.data[0] if insert_res.data else None
                
            if event_db:
                await update_auto_record_status_for_events(calendar, [event_db])
                fresh_db_res = supabase.table("calendar_events").select("*").eq("id", event_db["id"]).execute()
                if fresh_db_res.data:
                    event_db = fresh_db_res.data[0]
                    await sync_event_bot_schedule(user_id, event_db)

async def process_calendar_update_background(calendar_id: str):
    calendar_res = supabase.table("calendars").select("*").eq("id", calendar_id).execute()
    if not calendar_res.data:
        return
    calendar = calendar_res.data[0]
    recall_id = calendar["recall_id"]
    
    async with get_recall_client() as client:
        try:
            resp = await client.get(f"/api/v2/calendars/{recall_id}/")
            if resp.status_code == 200:
                recall_calendar = resp.json()
                settings = (calendar.get("recall_data") or {}).get("settings") or {}
                recall_calendar["settings"] = settings
                
                supabase.table("calendars").update({
                    "recall_data": recall_calendar
                }).eq("id", calendar_id).execute()
                print(f"Updated calendar {calendar_id} with fresh Recall data.")
        except Exception as e:
            print(f"Failed to update calendar metadata from Recall: {e}")

async def re_evaluate_calendar_events_background(calendar_id: str):
    calendar_res = supabase.table("calendars").select("*").eq("id", calendar_id).execute()
    if not calendar_res.data:
        return
    calendar = calendar_res.data[0]
    user_id = calendar["user_id"]
    
    events_res = supabase.table("calendar_events").select("*").eq("calendar_id", calendar_id).execute()
    events = events_res.data
    
    future_events = []
    for event in events:
        recall_data = event.get("recall_data") or {}
        end_time_str = recall_data.get("end_time")
        if end_time_str:
            try:
                clean_end_time = end_time_str.replace("Z", "+00:00")
                end_time = datetime.fromisoformat(clean_end_time)
                if end_time >= datetime.now(timezone.utc):
                    future_events.append(event)
            except Exception:
                pass
                
    if future_events:
        await update_auto_record_status_for_events(calendar, future_events)
        for event in future_events:
            fresh_res = supabase.table("calendar_events").select("*").eq("id", event["id"]).execute()
            if fresh_res.data:
                await sync_event_bot_schedule(user_id, fresh_res.data[0])


# ===== REQUEST SCHEMAS =====

class UpdateCalendarSettingsRequest(BaseModel):
    auto_record_external_events: Optional[bool] = None
    auto_record_only_confirmed_events: Optional[bool] = None

class UpdateEventSettingsRequest(BaseModel):
    should_record_manual: Optional[bool] = None


# ===== API ROUTE ENDPOINTS =====

@app.post("/webhooks/recall-calendar-updates")
async def recall_webhook(request: Request, background_tasks: BackgroundTasks):
    data = await request.json()
    print(f"INCOMING WEBHOOK from Recall: {data}")
    
    if supabase:
        try:
            supabase.table("calendar_webhooks").insert({
                "data": data
            }).execute()
            print("Successfully saved webhook to Supabase!")
        except Exception as e:
            print(f"CRITICAL Error saving webhook to Supabase: {e}")
    else:
        print("ERROR: Supabase client not initialized, cannot save webhook.")
        
    event = data.get("event")
    payload = data.get("data", {})
    recall_id = payload.get("calendar_id")
    
    if recall_id:
        if event == "calendar.update":
            calendar_res = supabase.table("calendars").select("id").eq("recall_id", recall_id).execute()
            if calendar_res.data:
                calendar_id = calendar_res.data[0]["id"]
                background_tasks.add_task(process_calendar_update_background, calendar_id)
        elif event == "calendar.sync_events":
            last_updated_ts = payload.get("last_updated_ts")
            background_tasks.add_task(process_calendar_sync_background, recall_id, last_updated_ts)
            
    return {"status": "ok"}

@app.get("/integrations/calendar")
async def get_calendars(request: Request):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
        
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase not initialized")
        
    res = supabase.table("calendars").select("*").eq("user_id", user_id).execute()
    return res.data

@app.get("/integrations/calendar/{calendar_id}/events")
async def get_calendar_events(calendar_id: str, request: Request):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
        
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase not initialized")
        
    cal_res = supabase.table("calendars").select("user_id").eq("id", calendar_id).execute()
    if not cal_res.data or cal_res.data[0]["user_id"] != user_id:
        raise HTTPException(status_code=403, detail="Forbidden: calendar does not belong to user")
        
    res = supabase.table("calendar_events").select("*").eq("calendar_id", calendar_id).execute()
    raw_events = res.data
    events = []
    for e in raw_events:
        recall_data = e.get("recall_data") or {}
        if recall_data.get("meeting_url"):
            events.append(e)
            
    def get_start_time(e):
        recall_data = e.get("recall_data") or {}
        return recall_data.get("start_time", "")
    events.sort(key=get_start_time)
    return events

@app.delete("/integrations/calendar/{calendar_id}")
async def delete_calendar(calendar_id: str, request: Request):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
        
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase not initialized")
        
    cal_res = supabase.table("calendars").select("*").eq("id", calendar_id).execute()
    if not cal_res.data or cal_res.data[0]["user_id"] != user_id:
        raise HTTPException(status_code=403, detail="Forbidden")
        
    calendar = cal_res.data[0]
    recall_id = calendar["recall_id"]
    
    # 1. Delete bot schedules for all events of this calendar
    events_res = supabase.table("calendar_events").select("recall_id").eq("calendar_id", calendar_id).execute()
    for event in events_res.data:
        try:
            await delete_bot_for_event(event["recall_id"])
        except Exception as e:
            print(f"Error deleting bot for event {event['recall_id']} during calendar deletion: {e}")
        
    # 2. Delete calendar in Recall.ai
    async with get_recall_client() as client:
        try:
            resp = await client.delete(f"/api/v2/calendars/{recall_id}/")
            if resp.status_code not in [200, 204, 404]:
                print(f"Recall calendar delete status: {resp.status_code} - {resp.text}")
        except Exception as e:
            print(f"Failed to delete calendar from Recall: {e}")
            
    # 3. Delete from Supabase
    try:
        supabase.table("calendar_events").delete().eq("calendar_id", calendar_id).execute()
        supabase.table("calendars").delete().eq("id", calendar_id).execute()
    except Exception as e:
        print(f"Error deleting calendar records from Supabase: {e}")
        raise HTTPException(status_code=500, detail=f"Database error: {e}")
    
    return {"status": "ok"}


@app.patch("/integrations/calendar/{calendar_id}")
async def update_calendar_settings(
    calendar_id: str,
    req_body: UpdateCalendarSettingsRequest,
    request: Request,
    background_tasks: BackgroundTasks
):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
        
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase not initialized")
        
    cal_res = supabase.table("calendars").select("*").eq("id", calendar_id).execute()
    if not cal_res.data or cal_res.data[0]["user_id"] != user_id:
        raise HTTPException(status_code=403, detail="Forbidden")
        
    calendar = cal_res.data[0]
    recall_data = calendar.get("recall_data") or {}
    settings = recall_data.get("settings") or {}
    
    if req_body.auto_record_external_events is not None:
        settings["auto_record_external_events"] = req_body.auto_record_external_events
    if req_body.auto_record_only_confirmed_events is not None:
        settings["auto_record_only_confirmed_events"] = req_body.auto_record_only_confirmed_events
        
    recall_data["settings"] = settings
    
    supabase.table("calendars").update({"recall_data": recall_data}).eq("id", calendar_id).execute()
    background_tasks.add_task(re_evaluate_calendar_events_background, calendar_id)
    
    return {"status": "ok", "settings": settings}

@app.patch("/integrations/calendar/event/{event_recall_id}")
async def update_event_settings(
    event_recall_id: str,
    req_body: UpdateEventSettingsRequest,
    request: Request,
    background_tasks: BackgroundTasks
):
    user_id = get_user_id_from_token(request.headers.get("Authorization"))
    if not user_id:
        user_id = DEFAULT_TEST_USER_ID
        
    if not supabase:
        raise HTTPException(status_code=500, detail="Supabase not initialized")
        
    event_res = supabase.table("calendar_events").select("*").eq("recall_id", event_recall_id).execute()
    if not event_res.data:
        raise HTTPException(status_code=404, detail="Event not found")
    event = event_res.data[0]
    calendar_id = event["calendar_id"]
    
    cal_res = supabase.table("calendars").select("user_id").eq("id", calendar_id).execute()
    if not cal_res.data or cal_res.data[0]["user_id"] != user_id:
        raise HTTPException(status_code=403, detail="Forbidden")
        
    recall_data = event.get("recall_data") or {}
    settings = recall_data.get("settings") or {}
    settings["should_record_manual"] = req_body.should_record_manual
    recall_data["settings"] = settings
    
    supabase.table("calendar_events").update({"recall_data": recall_data}).eq("id", event["id"]).execute()
    
    fresh_event_res = supabase.table("calendar_events").select("*").eq("id", event["id"]).execute()
    if fresh_event_res.data:
        background_tasks.add_task(sync_event_bot_schedule, user_id, fresh_event_res.data[0])
        
    return {"status": "ok", "settings": settings}

