import os
import uuid
from fastapi import FastAPI, UploadFile, File, Form, Depends, HTTPException
from fastapi.security import OAuth2PasswordRequestForm
from fastapi.responses import JSONResponse, FileResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.exc import IntegrityError
import msal
from fastapi import Request

from .celery_app import celery
from .tasks import convert_ppt_task
from .db import init_db, get_session
from .models import User
from .auth import (
    get_password_hash,
    verify_password,
    create_access_token,
    authenticate_user,
    get_current_user,
    encrypt_secret,
)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
UPLOADS = os.path.join(ROOT, "uploads")
os.makedirs(UPLOADS, exist_ok=True)

app = FastAPI(title="Generate Movie Web (MVP)")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# MSAL / Entra configuration pulled from environment
AZURE_CLIENT_ID = os.environ.get('AZURE_CLIENT_ID')
AZURE_CLIENT_SECRET = os.environ.get('AZURE_CLIENT_SECRET')
AZURE_TENANT_ID = os.environ.get('AZURE_TENANT_ID')
AZURE_REDIRECT_URI = os.environ.get('AZURE_REDIRECT_URI', 'http://localhost:8000/auth/callback')
AZURE_SCOPES = os.environ.get('AZURE_SCOPES', 'openid profile offline_access https://cognitiveservices.azure.com/.default').split()


def _build_msal_app():
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    return msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET,
    )


# Initialize the database on startup
@app.on_event("startup")
def on_startup():
    init_db()


@app.post('/register')
def register(username: str = Form(...), password: str = Form(...)):
    hashed = get_password_hash(password)
    user = User(username=username, hashed_password=hashed)
    try:
        with get_session() as session:
            session.add(user)
            session.commit()
            session.refresh(user)
            return {"id": user.id, "username": user.username}
    except IntegrityError:
        raise HTTPException(status_code=400, detail="Username already exists")


@app.post('/token')
def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends()):
    user = authenticate_user(form_data.username, form_data.password)
    if not user:
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    access_token = create_access_token(data={"sub": user.username})
    return {"access_token": access_token, "token_type": "bearer"}


@app.post('/credentials')
def set_credentials(endpoint: str = Form(...), key: str = Form(...), current_user: User = Depends(get_current_user)):
    # Encrypt and store Foundry credentials for the current user
    enc_endpoint = encrypt_secret(endpoint)
    enc_key = encrypt_secret(key)
    with get_session() as session:
        db_user = session.get(User, current_user.id)
        db_user.foundry_endpoint_encrypted = enc_endpoint
        db_user.foundry_key_encrypted = enc_key
        session.add(db_user)
        session.commit()
    return {"status": "saved"}


@app.post('/upload')
async def upload(ppt: UploadFile = File(...), voice_name: str = Form(None), current_user: User = Depends(get_current_user)):
    """Upload a PPTX and start a conversion job. Requires authentication."""
    filename = f"{uuid.uuid4().hex}_{ppt.filename}"
    dest = os.path.join(UPLOADS, filename)
    with open(dest, 'wb') as f:
        f.write(await ppt.read())

    task = convert_ppt_task.delay(dest, user_id=current_user.id, voice_name=voice_name)
    return {"job_id": task.id}


@app.get('/status/{job_id}')
def status(job_id: str):
    res = celery.AsyncResult(job_id)
    info = res.info if res.info else None
    return {"job_id": job_id, "state": res.state, "info": info}


@app.get('/result/{job_id}')
def result(job_id: str):
    res = celery.AsyncResult(job_id)
    if res.state == 'SUCCESS':
        info = res.result
        output_path = info.get('output_path') if isinstance(info, dict) else None
        if output_path and os.path.exists(output_path):
            return FileResponse(output_path, media_type='video/mp4', filename=os.path.basename(output_path))
        return JSONResponse({"error": "output missing", "detail": info}, status_code=404)
    else:
        return JSONResponse({"job_id": job_id, "state": res.state, "info": str(res.info)}, status_code=202)


@app.get('/login')
def login():
    app_msal = _build_msal_app()
    auth_url = app_msal.get_authorization_request_url(AZURE_SCOPES, redirect_uri=AZURE_REDIRECT_URI)
    return RedirectResponse(auth_url)


@app.get('/auth/callback')
def auth_callback(request: Request):
    code = request.query_params.get('code')
    if not code:
        raise HTTPException(status_code=400, detail='Missing code in callback')

    # Build an MSAL app that uses a SerializableTokenCache so we can persist it
    cache = msal.SerializableTokenCache()
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    app_msal = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET,
        token_cache=cache,
    )

    result = app_msal.acquire_token_by_authorization_code(code, scopes=AZURE_SCOPES, redirect_uri=AZURE_REDIRECT_URI)

    if 'error' in result:
        raise HTTPException(status_code=400, detail=f"Auth error: {result.get('error_description') or result.get('error')}")

    # Extract a username from id_token claims if available
    id_claims = result.get('id_token_claims', {})
    username = id_claims.get('preferred_username') or id_claims.get('upn') or id_claims.get('email') or id_claims.get('oid')
    if not username:
        raise HTTPException(status_code=400, detail='Could not determine username from id token')

    refresh_token = result.get('refresh_token')
    if not refresh_token:
        raise HTTPException(status_code=400, detail='No refresh token returned; ensure offline_access scope is granted')

    # Create or update local user and store encrypted refresh token + encrypted MSAL token cache
    from .auth import encrypt_secret
    with get_session() as session:
        user = session.query(User).filter(User.username == username).first()
        if not user:
            # Create local user record with a random password (they will authenticate via Entra)
            user = User(username=username, hashed_password=get_password_hash(uuid.uuid4().hex))
            session.add(user)
            session.commit()
            session.refresh(user)

        user.refresh_token_encrypted = encrypt_secret(refresh_token)
        # Persist serialized token cache as encrypted string so workers can call acquire_token_silent
        try:
            cache_serialized = cache.serialize()
            if cache_serialized:
                user.msal_token_cache_encrypted = encrypt_secret(cache_serialized)
        except Exception:
            # Non-fatal: keep refresh token fallback
            pass

        session.add(user)
        session.commit()

    # Issue a local JWT so the user can call API endpoints immediately
    token = create_access_token({"sub": username})
    return {"access_token": token, "token_type": "bearer"}
