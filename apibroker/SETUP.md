# Setup Instructions for NetDocuments Proxy API

## Prerequisites

- .NET 8 SDK
- SQL Server (LocalDB for development, SQL Server for production)
- Visual Studio 2022 or VS Code
- IIS (for production deployment)

## Development Setup

1. **Clone and Navigate**
   ```bash
   cd C:\dev\git\TN-OfficeDebug\apibroker
   ```

2. **Restore Dependencies**
   ```bash
   dotnet restore
   ```

3. **Configure User Secrets**
   ```bash
   dotnet user-secrets set "NetDocuments:ClientSecret" "iJ5RGF9eXXr86CFzl5kGbqEfT2RUIZf7GhXI8eHld1Km4gPY"
   dotnet user-secrets set "Jwt:Key" "your-256-bit-secret-key-here-make-it-very-long-and-secure"
   ```

4. **Update Database**
   ```bash
   dotnet ef database update
   ```

5. **Run Development Server**
   ```bash
   dotnet run
   ```

6. **Test API**
   - Navigate to `https://localhost:7001`
   - View Swagger documentation at `https://localhost:7001/swagger`

## Production Deployment

### IIS Setup

1. **Publish Application**
   ```bash
   dotnet publish -c Release -o publish
   ```

2. **Create IIS Application**
   - Create new application in IIS Manager
   - Set physical path to published folder
   - Configure application pool for .NET Core
   - Enable HTTPS and install SSL certificate

3. **Update Configuration**
   ```json
   {
     "ConnectionStrings": {
       "DefaultConnection": "Server=your-sql-server;Database=NetDocumentsProxy;Integrated Security=true;"
     },
     "NetDocuments": {
       "RedirectUri": "https://your-domain.com/api/auth/callback"
     }
   }
   ```

### Environment Variables (Alternative to User Secrets)

Set these in IIS application settings or web.config:
- `NetDocuments__ClientSecret`
- `Jwt__Key`

## Office Add-in Integration

### Update Office Add-in Configuration

1. **Update CORS Origin**
   In `appsettings.json`, update the CORS policy:
   ```json
   "AllowedOrigins": ["https://tools.bighand.services"]
   ```

2. **Update Office Add-in JavaScript**
   Replace the NetDocuments authentication calls in your Office add-in:

   ```javascript
   // Replace the old insecure authentication with this:
   async function authenticateWithNetDocuments(userId, email) {
     try {
       // Step 1: Get authorization URL
       const authStartResponse = await fetch(`${API_BASE_URL}/api/auth/start?userId=${encodeURIComponent(userId)}&email=${encodeURIComponent(email)}`);
       const { authUrl, state } = await authStartResponse.json();
       
       // Step 2: Open NetDocuments OAuth in popup
       Office.context.ui.displayDialogAsync(authUrl, {height: 60, width: 60}, (result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
           const dialog = result.value;
           
           dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
             const callbackData = JSON.parse(arg.message);
             dialog.close();
             
             // Step 3: Exchange authorization code for JWT token
             const tokenResponse = await fetch(`${API_BASE_URL}/api/auth/callback`, {
               method: 'POST',
               headers: { 'Content-Type': 'application/json' },
               body: JSON.stringify({
                 AuthorizationCode: callbackData.code,
                 State: callbackData.state
               })
             });
             
             const tokenData = await tokenResponse.json();
             if (tokenData.success) {
               // Store JWT token for API calls
               sessionStorage.setItem('api-token', tokenData.accessToken);
               console.log('Authentication successful!');
             }
           });
         }
       });
     } catch (error) {
       console.error('Authentication failed:', error);
     }
   }
   
   // Updated API call function
   async function getNetDocumentsDocument(documentId) {
     const token = sessionStorage.getItem('api-token');
     if (!token) {
       throw new Error('Not authenticated');
     }
     
     const response = await fetch(`${API_BASE_URL}/api/netdocuments/document/${documentId}`, {
       headers: {
         'Authorization': `Bearer ${token}`,
         'Content-Type': 'application/json'
       }
     });
     
     if (!response.ok) {
       throw new Error(`API call failed: ${response.status}`);
     }
     
     return await response.json();
   }
   ```

## Testing

### Test Authentication Flow
1. Start the API server
2. Navigate to `/api/auth/start?userId=test@bighand.com&email=test@bighand.com`
3. Complete OAuth flow with NetDocuments
4. Use returned JWT token for API calls

### Test Document Retrieval
```bash
curl -H "Authorization: Bearer YOUR_JWT_TOKEN" \
     https://localhost:7001/api/netdocuments/document/4124-1017-9932
```

## Monitoring

- **Health Checks**: `/health`
- **Logs**: Check `logs/` directory for application logs
- **Database**: Monitor `ApiLogs` table for API usage
- **Performance**: Monitor `UserTokens` table for authentication patterns

## Security Checklist

- [ ] JWT signing key is secure and unique
- [ ] NetDocuments client secret is stored securely
- [ ] HTTPS is enabled in production
- [ ] CORS is configured for specific origins only
- [ ] Rate limiting is enabled
- [ ] Database connection uses least privilege access
- [ ] Application runs under dedicated service account
- [ ] Logs are monitored for security events