using Microsoft.EntityFrameworkCore;
using NetDocumentsProxy.Data;

namespace NetDocumentsProxy.Services;

public class TokenRefreshService : BackgroundService
{
    private readonly IServiceProvider _serviceProvider;
    private readonly ILogger<TokenRefreshService> _logger;
    private readonly TimeSpan _refreshInterval = TimeSpan.FromHours(1); // Check every hour

    public TokenRefreshService(IServiceProvider serviceProvider, ILogger<TokenRefreshService> logger)
    {
        _serviceProvider = serviceProvider;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Token refresh service started");

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await RefreshExpiredTokensAsync();
                await Task.Delay(_refreshInterval, stoppingToken);
            }
            catch (OperationCanceledException)
            {
                _logger.LogInformation("Token refresh service is stopping");
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in token refresh service");
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken); // Wait 5 minutes before retrying
            }
        }
    }

    private async Task RefreshExpiredTokensAsync()
    {
        using var scope = _serviceProvider.CreateScope();
        var context = scope.ServiceProvider.GetRequiredService<NetDocumentsDbContext>();
        var oauthService = scope.ServiceProvider.GetRequiredService<OAuthService>();

        // Find tokens that will expire in the next 30 minutes
        var expiringTokens = await context.UserTokens
            .Where(t => t.ExpiresAt <= DateTime.UtcNow.AddMinutes(30))
            .ToListAsync();

        if (expiringTokens.Any())
        {
            _logger.LogInformation("Found {Count} tokens to refresh", expiringTokens.Count);

            foreach (var token in expiringTokens)
            {
                try
                {
                    // The OAuthService.GetValidAccessTokenAsync method will handle the refresh
                    var refreshedToken = await oauthService.GetValidAccessTokenAsync(token.UserId);
                    
                    if (refreshedToken != null)
                    {
                        _logger.LogInformation("Successfully refreshed token for user {UserId}", token.UserId);
                    }
                    else
                    {
                        _logger.LogWarning("Failed to refresh token for user {UserId}", token.UserId);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error refreshing token for user {UserId}", token.UserId);
                }

                // Add a small delay between refresh attempts to avoid overwhelming the OAuth server
                await Task.Delay(TimeSpan.FromSeconds(2));
            }
        }
        else
        {
            _logger.LogDebug("No tokens require refresh at this time");
        }
    }
}