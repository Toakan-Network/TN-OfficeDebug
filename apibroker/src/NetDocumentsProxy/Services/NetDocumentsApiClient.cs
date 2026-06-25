using Microsoft.EntityFrameworkCore;
using NetDocumentsProxy.Data;
using NetDocumentsProxy.Models;
using System.Diagnostics;
using System.Text.Json;

namespace NetDocumentsProxy.Services;

public class NetDocumentsApiClient
{
    private readonly HttpClient _httpClient;
    private readonly IConfiguration _configuration;
    private readonly NetDocumentsDbContext _context;
    private readonly ILogger<NetDocumentsApiClient> _logger;

    public NetDocumentsApiClient(
        HttpClient httpClient,
        IConfiguration configuration,
        NetDocumentsDbContext context,
        ILogger<NetDocumentsApiClient> logger)
    {
        _httpClient = httpClient;
        _configuration = configuration;
        _context = context;
        _logger = logger;
    }

    public async Task<NetDocumentsDocument?> GetDocumentAsync(string documentId, string accessToken, string? userId = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var endpoint = $"v2/document/{documentId}";
        
        try
        {
            _logger.LogInformation("Fetching document {DocumentId} for user {UserId}", documentId, userId);

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            var response = await _httpClient.GetAsync(endpoint);
            stopwatch.Stop();

            var responseContent = await response.Content.ReadAsStringAsync();

            // Log API call
            await LogApiCallAsync(userId, endpoint, "GET", (int)response.StatusCode, 
                stopwatch.ElapsedMilliseconds, null, responseContent, 
                response.IsSuccessStatusCode ? null : $"HTTP {response.StatusCode}");

            if (response.IsSuccessStatusCode)
            {
                var document = JsonSerializer.Deserialize<NetDocumentsDocument>(responseContent, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                _logger.LogInformation("Successfully retrieved document {DocumentId}", documentId);
                return document;
            }
            else
            {
                _logger.LogWarning("Failed to retrieve document {DocumentId}. Status: {StatusCode}, Response: {Response}", 
                    documentId, response.StatusCode, responseContent);
                return null;
            }
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex, "Error retrieving document {DocumentId}", documentId);
            
            await LogApiCallAsync(userId, endpoint, "GET", 500, 
                stopwatch.ElapsedMilliseconds, null, null, ex.Message);
            
            return null;
        }
    }

    public async Task<bool> TestConnectionAsync(string accessToken, string? userId = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var endpoint = "v2/repository";
        
        try
        {
            _logger.LogInformation("Testing NetDocuments API connection for user {UserId}", userId);

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            var response = await _httpClient.GetAsync(endpoint);
            stopwatch.Stop();

            var responseContent = await response.Content.ReadAsStringAsync();

            // Log API call
            await LogApiCallAsync(userId, endpoint, "GET", (int)response.StatusCode, 
                stopwatch.ElapsedMilliseconds, null, responseContent, 
                response.IsSuccessStatusCode ? null : $"HTTP {response.StatusCode}");

            if (response.IsSuccessStatusCode)
            {
                _logger.LogInformation("NetDocuments API connection test successful for user {UserId}", userId);
                return true;
            }
            else
            {
                _logger.LogWarning("NetDocuments API connection test failed for user {UserId}. Status: {StatusCode}", 
                    userId, response.StatusCode);
                return false;
            }
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex, "NetDocuments API connection test error for user {UserId}", userId);
            
            await LogApiCallAsync(userId, endpoint, "GET", 500, 
                stopwatch.ElapsedMilliseconds, null, null, ex.Message);
            
            return false;
        }
    }

    private async Task LogApiCallAsync(string? userId, string endpoint, string method, int statusCode, 
        long responseTime, string? requestBody, string? responseBody, string? errorMessage)
    {
        try
        {
            var apiLog = new ApiLog
            {
                UserId = userId,
                Endpoint = endpoint,
                Method = method,
                StatusCode = statusCode,
                ResponseTime = responseTime,
                RequestedAt = DateTime.UtcNow,
                RequestBody = requestBody,
                ResponseBody = responseBody?.Length > 1000 ? responseBody.Substring(0, 1000) + "..." : responseBody,
                ErrorMessage = errorMessage
            };

            _context.ApiLogs.Add(apiLog);
            await _context.SaveChangesAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to log API call");
        }
    }
}