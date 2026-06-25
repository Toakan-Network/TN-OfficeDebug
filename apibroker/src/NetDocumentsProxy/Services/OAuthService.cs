using Microsoft.EntityFrameworkCore;
using NetDocumentsProxy.Data;
using NetDocumentsProxy.Models;
using System.Text.Json;

namespace NetDocumentsProxy.Services;

public class OAuthService
{
    private readonly IConfiguration _configuration;
    private readonly HttpClient _httpClient;
    private readonly NetDocumentsDbContext _context;
    private readonly ILogger<OAuthService> _logger;

    public OAuthService(
        IConfiguration configuration,
        HttpClient httpClient,
        NetDocumentsDbContext context,
        ILogger<OAuthService> logger)
    {
        _configuration = configuration;
        _httpClient = httpClient;
        _context = context;
        _logger = logger;
    }

    public string GetAuthorizationUrl(string state)
    {
        var clientId = _configuration["NetDocuments:ClientId"];
        var redirectUri = _configuration["NetDocuments:RedirectUri"];
        var oauthUrl = _configuration["NetDocuments:OAuthUrl"];

        var authUrl = $"{oauthUrl}/v2/oauth2/auth?" +
                     $"response_type=code&" +
                     $"client_id={Uri.EscapeDataString(clientId!)}&" +
                     $"redirect_uri={Uri.EscapeDataString(redirectUri!)}&" +
                     $"scope=read%20write&" +
                     $"state={Uri.EscapeDataString(state)}";

        _logger.LogInformation("Generated authorization URL for state: {State}", state);
        return authUrl;
    }

    public async Task<AuthResponse> ExchangeCodeForTokenAsync(string code, string userId)
    {
        try
        {
            var tokenEndpoint = $"{_configuration["NetDocuments:OAuthUrl"]}/v2/oauth2/token";
            
            var tokenRequest = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "authorization_code"),
                new KeyValuePair<string, string>("code", code),
                new KeyValuePair<string, string>("client_id", _configuration["NetDocuments:ClientId"]!),
                new KeyValuePair<string, string>("client_secret", _configuration["NetDocuments:ClientSecret"]!),
                new KeyValuePair<string, string>("redirect_uri", _configuration["NetDocuments:RedirectUri"]!)
            });

            _logger.LogInformation("Exchanging authorization code for access token for user: {UserId}", userId);

            var response = await _httpClient.PostAsync(tokenEndpoint, tokenRequest);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                var tokenResponse = JsonSerializer.Deserialize<OAuthTokenResponse>(responseContent);
                if (tokenResponse != null)
                {
                    await StoreUserTokenAsync(userId, tokenResponse);
                    
                    _logger.LogInformation("Successfully obtained and stored tokens for user: {UserId}", userId);
                    
                    return new AuthResponse
                    {
                        Success = true,
                        AccessToken = tokenResponse.access_token,
                        ExpiresAt = DateTime.UtcNow.AddSeconds(tokenResponse.expires_in)
                    };
                }
            }

            _logger.LogWarning("Token exchange failed for user {UserId}. Status: {StatusCode}, Response: {Response}", 
                userId, response.StatusCode, responseContent);

            return new AuthResponse
            {
                Success = false,
                Error = $"Token exchange failed: {response.StatusCode}"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during token exchange for user: {UserId}", userId);
            return new AuthResponse
            {
                Success = false,
                Error = "Internal server error during token exchange"
            };
        }
    }

    public async Task<string?> GetValidAccessTokenAsync(string userId)
    {
        var userToken = await _context.UserTokens
            .FirstOrDefaultAsync(t => t.UserId == userId);

        if (userToken == null)
        {
            _logger.LogWarning("No token found for user: {UserId}", userId);
            return null;
        }

        // Check if token is still valid (with 5 minute buffer)
        if (userToken.ExpiresAt > DateTime.UtcNow.AddMinutes(5))
        {
            _logger.LogDebug("Using existing valid token for user: {UserId}", userId);
            return userToken.AccessToken;
        }

        // Try to refresh the token
        _logger.LogInformation("Token expired for user {UserId}, attempting refresh", userId);
        var refreshed = await RefreshTokenAsync(userToken);
        
        return refreshed ? userToken.AccessToken : null;
    }

    private async Task<bool> RefreshTokenAsync(UserToken userToken)
    {
        try
        {
            var tokenEndpoint = $"{_configuration["NetDocuments:OAuthUrl"]}/v2/oauth2/token";
            
            var refreshRequest = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "refresh_token"),
                new KeyValuePair<string, string>("refresh_token", userToken.RefreshToken),
                new KeyValuePair<string, string>("client_id", _configuration["NetDocuments:ClientId"]!),
                new KeyValuePair<string, string>("client_secret", _configuration["NetDocuments:ClientSecret"]!)
            });

            var response = await _httpClient.PostAsync(tokenEndpoint, refreshRequest);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                var tokenResponse = JsonSerializer.Deserialize<OAuthTokenResponse>(responseContent);
                if (tokenResponse != null)
                {
                    userToken.AccessToken = tokenResponse.access_token;
                    userToken.RefreshToken = tokenResponse.refresh_token;
                    userToken.ExpiresAt = DateTime.UtcNow.AddSeconds(tokenResponse.expires_in);
                    userToken.UpdatedAt = DateTime.UtcNow;

                    _context.UserTokens.Update(userToken);
                    await _context.SaveChangesAsync();

                    _logger.LogInformation("Successfully refreshed token for user: {UserId}", userToken.UserId);
                    return true;
                }
            }

            _logger.LogWarning("Token refresh failed for user {UserId}. Status: {StatusCode}", 
                userToken.UserId, response.StatusCode);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during token refresh for user: {UserId}", userToken.UserId);
            return false;
        }
    }

    private async Task StoreUserTokenAsync(string userId, OAuthTokenResponse tokenResponse)
    {
        var existingToken = await _context.UserTokens
            .FirstOrDefaultAsync(t => t.UserId == userId);

        var now = DateTime.UtcNow;
        var expiresAt = now.AddSeconds(tokenResponse.expires_in);

        if (existingToken != null)
        {
            existingToken.AccessToken = tokenResponse.access_token;
            existingToken.RefreshToken = tokenResponse.refresh_token;
            existingToken.ExpiresAt = expiresAt;
            existingToken.UpdatedAt = now;
            _context.UserTokens.Update(existingToken);
        }
        else
        {
            var newToken = new UserToken
            {
                UserId = userId,
                AccessToken = tokenResponse.access_token,
                RefreshToken = tokenResponse.refresh_token,
                ExpiresAt = expiresAt,
                CreatedAt = now,
                UpdatedAt = now
            };
            _context.UserTokens.Add(newToken);
        }

        await _context.SaveChangesAsync();
    }
}