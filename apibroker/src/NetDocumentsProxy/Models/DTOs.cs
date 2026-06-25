namespace NetDocumentsProxy.Models;

public class OAuthTokenResponse
{
    public string access_token { get; set; } = string.Empty;
    public string refresh_token { get; set; } = string.Empty;
    public string token_type { get; set; } = "Bearer";
    public int expires_in { get; set; }
    public string scope { get; set; } = string.Empty;
}

public class NetDocumentsDocument
{
    public string id { get; set; } = string.Empty;
    public string name { get; set; } = string.Empty;
    public string extension { get; set; } = string.Empty;
    public long size { get; set; }
    public DateTime created { get; set; }
    public DateTime modified { get; set; }
    public string? author { get; set; }
    public string? version { get; set; }
    public string? status { get; set; }
    public DocumentAttribute[]? attributes { get; set; }
}

public class DocumentAttribute
{
    public string name { get; set; } = string.Empty;
    public string value { get; set; } = string.Empty;
}

public class AuthRequest
{
    public string AuthorizationCode { get; set; } = string.Empty;
    public string State { get; set; } = string.Empty;
}

public class AuthResponse
{
    public string AccessToken { get; set; } = string.Empty;
    public DateTime ExpiresAt { get; set; }
    public bool Success { get; set; }
    public string? Error { get; set; }
}