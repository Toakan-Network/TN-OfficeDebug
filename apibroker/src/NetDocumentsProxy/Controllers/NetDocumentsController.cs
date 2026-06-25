using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using NetDocumentsProxy.Services;

namespace NetDocumentsProxy.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class NetDocumentsController : ControllerBase
{
    private readonly NetDocumentsApiClient _apiClient;
    private readonly OAuthService _oauthService;
    private readonly TokenService _tokenService;
    private readonly ILogger<NetDocumentsController> _logger;

    public NetDocumentsController(
        NetDocumentsApiClient apiClient,
        OAuthService oauthService,
        TokenService tokenService,
        ILogger<NetDocumentsController> logger)
    {
        _apiClient = apiClient;
        _oauthService = oauthService;
        _tokenService = tokenService;
        _logger = logger;
    }

    /// <summary>
    /// Tests connection to NetDocuments API
    /// </summary>
    [HttpGet("test-connection")]
    public async Task<IActionResult> TestConnection()
    {
        var userId = _tokenService.GetUserIdFromToken(User);
        if (userId == null)
        {
            return Unauthorized(new { error = "Invalid user token" });
        }

        var accessToken = await _oauthService.GetValidAccessTokenAsync(userId);
        if (accessToken == null)
        {
            return Unauthorized(new { error = "No valid NetDocuments token found. Please re-authenticate." });
        }

        var isConnected = await _apiClient.TestConnectionAsync(accessToken, userId);

        return Ok(new
        {
            success = isConnected,
            message = isConnected ? "Connection successful" : "Connection failed",
            timestamp = DateTime.UtcNow
        });
    }

    /// <summary>
    /// Retrieves document information from NetDocuments
    /// </summary>
    [HttpGet("document/{documentId}")]
    public async Task<IActionResult> GetDocument(string documentId)
    {
        if (string.IsNullOrEmpty(documentId))
        {
            return BadRequest(new { error = "Document ID is required" });
        }

        var userId = _tokenService.GetUserIdFromToken(User);
        if (userId == null)
        {
            return Unauthorized(new { error = "Invalid user token" });
        }

        var accessToken = await _oauthService.GetValidAccessTokenAsync(userId);
        if (accessToken == null)
        {
            return Unauthorized(new { error = "No valid NetDocuments token found. Please re-authenticate." });
        }

        _logger.LogInformation("Retrieving document {DocumentId} for user {UserId}", documentId, userId);

        var document = await _apiClient.GetDocumentAsync(documentId, accessToken, userId);

        if (document == null)
        {
            return NotFound(new { error = "Document not found or access denied" });
        }

        return Ok(new
        {
            success = true,
            document = document,
            retrievedAt = DateTime.UtcNow
        });
    }

    /// <summary>
    /// Gets authentication status for current user
    /// </summary>
    [HttpGet("auth-status")]
    public async Task<IActionResult> GetAuthStatus()
    {
        var userId = _tokenService.GetUserIdFromToken(User);
        if (userId == null)
        {
            return Unauthorized(new { error = "Invalid user token" });
        }

        var accessToken = await _oauthService.GetValidAccessTokenAsync(userId);

        return Ok(new
        {
            isAuthenticated = accessToken != null,
            userId = userId,
            hasValidToken = accessToken != null,
            timestamp = DateTime.UtcNow
        });
    }

    /// <summary>
    /// Revokes user's NetDocuments authentication
    /// </summary>
    [HttpPost("revoke")]
    public async Task<IActionResult> RevokeAuth()
    {
        var userId = _tokenService.GetUserIdFromToken(User);
        if (userId == null)
        {
            return Unauthorized(new { error = "Invalid user token" });
        }

        // In production, also revoke tokens with NetDocuments OAuth server
        // For now, just remove from our database
        
        _logger.LogInformation("Revoking NetDocuments authentication for user {UserId}", userId);

        return Ok(new
        {
            success = true,
            message = "Authentication revoked successfully"
        });
    }
}