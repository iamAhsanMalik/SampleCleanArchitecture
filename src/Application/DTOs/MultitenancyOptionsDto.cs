namespace Application.DTOs;

public class MultitenancyOptionsDto
{
    public ICollection<AppTenantDto> Tenants { get; set; } = new List<AppTenantDto>();
}