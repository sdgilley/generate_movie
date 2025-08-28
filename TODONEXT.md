# Local Azure App Registration Setup Steps

1. Register an app in Azure AD (Azure Portal → Azure Active Directory → App registrations → New registration).
   - Name: any descriptive name (e.g. generate_movie-local).
   - Redirect URI (Web): set to `http://localhost:8000/auth/callback`.
   - Supported account types: pick tenant-only or multi‑tenant as needed.

2. After registration copy these values:
   - `AZURE_CLIENT_ID` = the Application (client) ID shown on the app’s Overview page.
   - `AZURE_TENANT_ID` = the Directory (tenant) ID shown on the app’s Overview page.

3. Create a client secret:
   - Portal → Certificates & secrets → New client secret → create and copy the *Value* immediately.
   - `AZURE_CLIENT_SECRET` = that secret value (you will not be able to view it again later).

4. Scopes / consent:
   - The app should request `offline_access openid profile` and the resource scope `https://cognitiveservices.azure.com/.default` at auth time (the code flow will request `offline_access` so you receive a refresh token).
   - If you want tenant-wide consent, an admin must grant consent in the Portal.

5. Put them into your `.env` (example):
   ```properties
   AZURE_CLIENT_ID=your-client-id-guid
   AZURE_CLIENT_SECRET=your-client-secret-value
   AZURE_TENANT_ID=your-tenant-id-guid
   AZURE_REDIRECT_URI=http://localhost:8000/auth/callback
   ```

6. Security note:
   - Treat `AZURE_CLIENT_SECRET` as a secret (use Key Vault or encrypted storage in production).
   - Copy the client secret value right after creation — it is shown only once.

7. After updating `.env`, restart your API and Celery worker for changes to take effect.
