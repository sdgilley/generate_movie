import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path
import requests
import msal

from .celery_app import celery

# Ensure utilities/ is importable
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
UTILS_DIR = os.path.join(ROOT, "utilities")
if os.path.isdir(UTILS_DIR):
    sys.path.insert(0, UTILS_DIR)

# Import DB and auth utilities for secure credential retrieval
from .db import get_session
from .models import User
from .auth import decrypt_secret, encrypt_secret

@celery.task(bind=True)
def convert_ppt_task(self, ppt_path, user_id=None, voice_name=None):
    """Celery task to convert an uploaded PPTX into an MP4 using existing utilities.

    This task creates an isolated temporary working directory, copies the uploaded
    pptx into it, sets environment variables for the job (POWERPOINT_FILE, SPEECH_KEY, ENDPOINT,
    VOICE_NAME) and then calls the existing `generate_from_slides` and
    `generate_with_azure_audio` utilities.

    If `user_id` is provided, the task will fetch the user's encrypted Foundry
    credentials (or stored refresh token) from the local database and decrypt them server-side before use.

    Returns a dict with status and output_path (absolute file path) on success.
    """
    job_dir = tempfile.mkdtemp(prefix="gm_job_")
    cwd = os.getcwd()

    try:
        # Copy uploaded pptx into job dir
        dest_pptx = os.path.join(job_dir, os.path.basename(ppt_path))
        shutil.copy(ppt_path, dest_pptx)

        # If a user_id was supplied, fetch encrypted credentials from DB and decrypt
        speech_key = None
        endpoint = None
        access_token = None
        if user_id:
            try:
                with get_session() as session:
                    user = session.get(User, user_id)
                    if user:
                        # If a refresh token exists, try to exchange it for an access token
                        if getattr(user, 'refresh_token_encrypted', None):
                            try:
                                refresh_token = decrypt_secret(user.refresh_token_encrypted)
                                tenant = os.environ.get('AZURE_TENANT_ID')
                                client_id = os.environ.get('AZURE_CLIENT_ID')
                                client_secret = os.environ.get('AZURE_CLIENT_SECRET')

                                # Prefer using persisted MSAL token cache to acquire tokens silently
                                if getattr(user, 'msal_token_cache_encrypted', None):
                                    try:
                                        cache = msal.SerializableTokenCache()
                                        cache.deserialize(decrypt_secret(user.msal_token_cache_encrypted))
                                        authority = f'https://login.microsoftonline.com/{tenant}'
                                        msal_app = msal.ConfidentialClientApplication(
                                            client_id,
                                            authority=authority,
                                            client_credential=client_secret,
                                            token_cache=cache,
                                        )

                                        accounts = msal_app.get_accounts()
                                        account = None
                                        for a in accounts:
                                            if a.get('username') and a.get('username').lower() == user.username.lower():
                                                account = a
                                                break

                                        if account:
                                            result = msal_app.acquire_token_silent(['https://cognitiveservices.azure.com/.default'], account=account)
                                            if result and 'access_token' in result:
                                                access_token = result.get('access_token')
                                                os.environ['AZURE_ACCESS_TOKEN'] = access_token
                                                # Persist updated cache if MSAL mutated it
                                                try:
                                                    serialized = cache.serialize()
                                                    if serialized:
                                                        user.msal_token_cache_encrypted = encrypt_secret(serialized)
                                                        session.add(user)
                                                        session.commit()
                                                except Exception as e:
                                                    self.update_state(state='PROGRESS', meta={'step': 'cache_persist_failed', 'error': str(e)})
                                            else:
                                                # Silent acquire failed - try refresh-token exchange using MSAL to populate cache
                                                result = msal_app.acquire_token_by_refresh_token(refresh_token, scopes=['https://cognitiveservices.azure.com/.default'])
                                                if result and 'access_token' in result:
                                                    access_token = result.get('access_token')
                                                    os.environ['AZURE_ACCESS_TOKEN'] = access_token
                                                    try:
                                                        serialized = cache.serialize()
                                                        if serialized:
                                                            user.msal_token_cache_encrypted = encrypt_secret(serialized)
                                                            session.add(user)
                                                            session.commit()
                                                    except Exception:
                                                        pass
                                        else:
                                            # No matching account in cache; try refresh token via MSAL
                                            result = msal_app.acquire_token_by_refresh_token(refresh_token, scopes=['https://cognitiveservices.azure.com/.default'])
                                            if result and 'access_token' in result:
                                                access_token = result.get('access_token')
                                                os.environ['AZURE_ACCESS_TOKEN'] = access_token
                                                try:
                                                    serialized = cache.serialize()
                                                    if serialized:
                                                        user.msal_token_cache_encrypted = encrypt_secret(serialized)
                                                        session.add(user)
                                                        session.commit()
                                                except Exception:
                                                    pass
                                    except Exception as e:
                                        # MSAL cache path failed; fallback to MSAL refresh token exchange below
                                        self.update_state(state='PROGRESS', meta={'step': 'msal_cache_failed', 'error': str(e)})
                                # If we get here and no access_token yet, try MSAL refresh-token exchange
                                if not access_token:
                                    try:
                                        authority = f'https://login.microsoftonline.com/{tenant}'
                                        msal_app = msal.ConfidentialClientApplication(
                                            client_id,
                                            authority=authority,
                                            client_credential=client_secret,
                                        )
                                        result = msal_app.acquire_token_by_refresh_token(refresh_token, scopes=['https://cognitiveservices.azure.com/.default'])
                                        if result and 'access_token' in result:
                                            access_token = result.get('access_token')
                                            os.environ['AZURE_ACCESS_TOKEN'] = access_token
                                    except Exception as e:
                                        # MSAL refresh failed; fall back to direct token endpoint POST
                                        try:
                                            token_url = f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'
                                            data = {
                                                'client_id': client_id,
                                                'client_secret': client_secret,
                                                'grant_type': 'refresh_token',
                                                'refresh_token': refresh_token,
                                                'scope': 'https://cognitiveservices.azure.com/.default'
                                            }
                                            resp = requests.post(token_url, data=data, timeout=30)
                                            if resp.ok:
                                                token_json = resp.json()
                                                access_token = token_json.get('access_token')
                                                if access_token:
                                                    os.environ['AZURE_ACCESS_TOKEN'] = access_token
                                            else:
                                                self.update_state(state='PROGRESS', meta={'step': 'token_exchange_failed', 'error': resp.text})
                                        except Exception as se:
                                            self.update_state(state='PROGRESS', meta={'step': 'token_exchange_failed', 'error': str(se)})
                            except Exception as e:
                                # failure to exchange token; fall back to stored key if present
                                self.update_state(state='PROGRESS', meta={'step': 'token_exchange_failed', 'error': str(e)})
                        # Backward compatibility: if the user stored a direct key/endpoint earlier
                        if user.foundry_key_encrypted:
                            speech_key = decrypt_secret(user.foundry_key_encrypted)
                        if user.foundry_endpoint_encrypted:
                            endpoint = decrypt_secret(user.foundry_endpoint_encrypted)
            except Exception as e:
                # Don't fail the job immediately on DB errors; fall back to environment
                self.update_state(state='PROGRESS', meta={'step': 'db_error', 'error': str(e)})

        # Set env vars before importing modules that read them at import-time
        os.environ['POWERPOINT_FILE'] = dest_pptx
        if speech_key:
            os.environ['SPEECH_KEY'] = speech_key
        if endpoint:
            os.environ['ENDPOINT'] = endpoint
        if voice_name:
            os.environ['VOICE_NAME'] = voice_name

        # Switch to job dir so utility lookups (exported_slides, movies/) are local
        os.chdir(job_dir)

        # Import utilities now that POWERPOINT_FILE is set and sys.path includes utilities/
        import generate_from_slides as gfs
        import generate_with_azure_audio as gwa

        # Export slides (use Python fallback which is cross-platform)
        self.update_state(state='PROGRESS', meta={'step': 'export_slides'})
        try:
            exported_ok = gfs.export_slides_python_fallback(dest_pptx, output_dir='exported_slides')
        except Exception as e:
            exported_ok = False
            self.update_state(state='PROGRESS', meta={'step': 'export_failed', 'error': str(e)})

        # Run the main generator
        self.update_state(state='PROGRESS', meta={'step': 'generate_video'})
        success = gwa.main()

        # Determine the output path
        output_rel = getattr(gwa, 'output_video_name', None)
        if output_rel and not os.path.isabs(output_rel):
            output_path = os.path.join(job_dir, output_rel)
        else:
            output_path = output_rel

        if success and output_path and os.path.exists(output_path):
            return {'status': 'SUCCESS', 'output_path': output_path}
        else:
            return {'status': 'FAILED', 'output_path': output_path}

    except Exception as exc:
        traceback.print_exc()
        return {'status': 'ERROR', 'error': str(exc)}

    finally:
        # Restore cwd (don't cleanup job_dir in MVP so users can retrieve files)
        os.chdir(cwd)
