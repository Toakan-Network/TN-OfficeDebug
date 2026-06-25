using Microsoft.AspNetCore.Mvc;
using NetDocumentsProxy.Models;
using NetDocumentsProxy.Services;
using System.Security.Cryptography;
using System.Text;

namespace NetDocumentsProxy.Controllers;

[ApiController]
[Route("api/[controller]")]
public class AuthController : ControllerBase
{
    private readonly OAuthService _oauthService;
    private readonly TokenService _tokenService;
    private readonly ILogger<AuthController> _logger;

    public AuthController(
        OAuthService oauthService,
        TokenService tokenService,
        ILogger<AuthController> logger)
    {
        _oauthService = oauthService;
        _tokenService = tokenService;
        _logger = logger;
    }

    /// <summary>
    /// Initiates OAuth flow by returning authorization URL
    /// </summary>
    [HttpGet("start")]
    public IActionResult StartAuth([FromQuery] string userId, [FromQuery] string email)
    {
        if (string.IsNullOrEmpty(userId) || string.IsNullOrEmpty(email))
        {
            return BadRequest(new { error = "userId and email are required" });
        }

        // Generate secure state parameter
        var state = GenerateSecureState();
        
        // Store state in cache/session for validation (simplified for demo)
        HttpContext.Session.SetString($"oauth_state_{userId}", state);
        HttpContext.Session.SetString($"oauth_email_{userId}", email);

        var authUrl = _oauthService.GetAuthorizationUrl(state);

        _logger.LogInformation("OAuth flow started for user {UserId}", userId);

        return Ok(new
        {
            authUrl = authUrl,
            state = state
        });
    }

    /// <summary>
    /// Handles OAuth callback and exchanges code for tokens
    /// </summary>
    [HttpPost("callback")]
    public async Task<IActionResult> OAuthCallback([FromBody] AuthRequest request)
    {
        if (string.IsNullOrEmpty(request.AuthorizationCode) || string.IsNullOrEmpty(request.State))
        {
            return BadRequest(new { error = "Authorization code and state are required" });
        }

        // Extract user ID from state (in production, use more secure state management)
        var userId = ExtractUserIdFromState(request.State);
        if (string.IsNullOrEmpty(userId))
        {
            return BadRequest(new { error = "Invalid state parameter" });
        }

        // Validate state
        var storedState = HttpContext.Session.GetString($"oauth_state_{userId}");
        if (storedState != request.State)
        {
            _logger.LogWarning("OAuth state mismatch for user {UserId}", userId);
            return BadRequest(new { error = "Invalid state parameter" });
        }

        var email = HttpContext.Session.GetString($"oauth_email_{userId}") ?? userId;

        // Exchange code for tokens
        var result = await _oauthService.ExchangeCodeForTokenAsync(request.AuthorizationCode, userId);

        if (result.Success)
        {
            // Generate JWT token for API access
            var jwtToken = _tokenService.GenerateJwtToken(userId, email);

            // Clean up session
            HttpContext.Session.Remove($"oauth_state_{userId}");
            HttpContext.Session.Remove($"oauth_email_{userId}");

            _logger.LogInformation("OAuth flow completed successfully for user {UserId}", userId);

            return Ok(new
            {
                success = true,
                accessToken = jwtToken,
                expiresAt = result.ExpiresAt
            });
        }

        _logger.LogWarning("OAuth flow failed for user {UserId}: {Error}", userId, result.Error);
        return BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Validates and refreshes JWT token
    /// </summary>
    [HttpPost("refresh")]
    public IActionResult RefreshToken([FromHeader(Name = "Authorization")] string? authorization)
    {
        if (string.IsNullOrEmpty(authorization) || !authorization.StartsWith("Bearer "))
        {
            return Unauthorized(new { error = "Valid bearer token required" });
        }

        var token = authorization.Substring("Bearer ".Length);
        var principal = _tokenService.ValidateToken(token);

        if (principal == null)
        {
            return Unauthorized(new { error = "Invalid token" });
        }

        var userId = _tokenService.GetUserIdFromToken(principal);
        var email = principal.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ?? userId!;

        // Generate new JWT token
        var newToken = _tokenService.GenerateJwtToken(userId!, email);

        return Ok(new
        {
            accessToken = newToken,
            expiresAt = DateTime.UtcNow.AddDays(7)
        });
    }

    private string GenerateSecureState()
    {
        using var rng = RandomNumberGenerator.Create();
        var bytes = new byte[32];
        rng.GetBytes(bytes);
        return Convert.ToBase64String(bytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
    }

    private string? ExtractUserIdFromState(string state)
    {
        // In production, implement proper state management with database/cache
        // For demo, we'll extract from session during callback
        foreach (var key in HttpContext.Session.Keys)
        {
            if (key.StartsWith("oauth_state_"))
            {
                var storedState = HttpContext.Session.GetString(key);
                if (storedState == state)
                {
                    return key.Substring("oauth_state_".Length);
                }
            }
        }
        return null;
    }
}