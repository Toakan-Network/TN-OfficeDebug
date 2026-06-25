using System.ComponentModel.DataAnnotations;

namespace NetDocumentsProxy.Models;

public class ApiLog
{
    public int Id { get; set; }
    
    [MaxLength(100)]
    public string? UserId { get; set; }
    
    [Required]
    [MaxLength(500)]
    public string Endpoint { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(10)]
    public string Method { get; set; } = string.Empty;
    
    public int StatusCode { get; set; }
    
    public DateTime RequestedAt { get; set; }
    
    public long ResponseTime { get; set; }
    
    public string? RequestBody { get; set; }
    
    public string? ResponseBody { get; set; }
    
    public string? ErrorMessage { get; set; }
}