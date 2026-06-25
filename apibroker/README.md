# NetDocuments Proxy API

This ASP.NET Core Web API serves as a secure proxy for NetDocuments OAuth authentication and API integration, specifically designed to work with the TN-OfficeDebug Office Add-in.

## Features

- **Secure OAuth 2.0 Authorization Code Flow**: No password collection, uses proper OAuth standards
- **CORS-Compliant API Proxy**: Handles NetDocuments API calls without CORS restrictions
- **JWT Authentication**: Secure token-based authentication for API access
- **Automatic Token Refresh**: Background service maintains valid NetDocuments tokens
- **Comprehensive Logging**: Full audit trail of API calls and authentication events
- **Rate Limiting**: Protection against API abuse
- **Health Checks**: Monitoring endpoints for production deployment

## Architecture

### Authentication Flow
1. Office Add-in requests authorization URL from `/api/auth/start`
2. User is redirected to NetDocuments OAuth login (secure popup)
3. NetDocuments redirects back to `/api/auth/callback` with authorization code
4. API exchanges code for NetDocuments access/refresh tokens
5. API returns JWT token for subsequent API calls

### API Proxy Flow
1. Office Add-in calls API endpoints with JWT token
2. API validates JWT and retrieves stored NetDocuments tokens
3. API makes authenticated calls to NetDocuments on behalf of user
4. API returns results to Office Add-in with proper CORS headers

## Configuration

### Required User Secrets
```bash
dotnet user-secrets set "NetDocuments:ClientSecret" "your-client-secret"
dotnet user-secrets set "Jwt:Key" "your-very-long-secure-jwt-signing-key"
```

### Database Setup
The application uses Entity Framework with SQL Server. Update the connection string in `appsettings.json` or `appsettings.Production.json`.

### IIS Deployment
1. Publish the application: `dotnet publish -c Release`
2. Copy published files to IIS application directory
3. Ensure application pool runs under .NET Core
4. Update connection string for production database
5. Configure SSL certificate for HTTPS

## API Endpoints

### Authentication
- `GET /api/auth/start` - Start OAuth flow
- `POST /api/auth/callback` - Handle OAuth callback
- `POST /api/auth/refresh` - Refresh JWT token

### NetDocuments Integration
- `GET /api/netdocuments/test-connection` - Test API connectivity
- `GET /api/netdocuments/document/{id}` - Get document information
- `GET /api/netdocuments/auth-status` - Check authentication status
- `POST /api/netdocuments/revoke` - Revoke authentication

### System
- `GET /health` - Health check endpoint
- `GET /swagger` - API documentation (development only)

## Security Considerations

- All endpoints use HTTPS in production
- JWT tokens have configurable expiration
- NetDocuments tokens are stored securely in database
- Rate limiting prevents API abuse
- Comprehensive logging for security monitoring
- CORS configured for specific Office Add-in origin

## Office Add-in Integration

Update your Office Add-in JavaScript to use this proxy:

```javascript
// Start OAuth flow
const authResponse = await fetch('https://your-api-domain/api/auth/start?userId=user@domain.com&email=user@domain.com');
const { authUrl } = await authResponse.json();

// Open popup for OAuth (use Office Dialog API)
Office.context.ui.displayDialogAsync(authUrl, {height: 60, width: 60}, (result) => {
    // Handle OAuth callback and get JWT token
});

// Use JWT token for API calls
const documentResponse = await fetch('https://your-api-domain/api/netdocuments/document/4124-1017-9932', {
    headers: {
        'Authorization': `Bearer ${jwtToken}`,
        'Content-Type': 'application/json'
    }
});
```