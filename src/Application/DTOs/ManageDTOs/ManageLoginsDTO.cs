namespace Application.DTOs.ManageDTOs;

public class ManageLoginsDTO<IdentityUserLoginInfo, IdentityAuthenticationScheme>
{
    public IList<IdentityUserLoginInfo>? CurrentLogins { get; set; }

    public IList<IdentityAuthenticationScheme>? OtherLogins { get; set; }
}
