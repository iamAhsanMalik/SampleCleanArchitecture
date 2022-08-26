namespace Application.DTOs;

public class AppTenantDto
{
    public string? Name { get; set; }
    public string[]? Hostnames { get; set; }
    public string? Theme { get; set; }
    public string? ConnectionString { get; set; }
}