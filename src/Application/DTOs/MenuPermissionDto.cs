namespace Application.DTOs;

public class MenuPermissionDto
{
    public string? CpUserID { get; set; }
    public int MenuID { get; set; }
    //public string cpUserType
    public string? ActionControls { get; set; }
    public string? PnlLeft { get; set; }
    public string? PnlRight { get; set; }
    public string? ErrMessage { get; set; }
    public string? SecurityStamp { get; set; }

}
public class UserPermissionDto
{
    public string? UserID { get; set; }
    public bool IsValid { get; set; }
    public string? Message { get; set; }
    public string? PinCode { get; set; }
    public List<UserMenuDto> Menus { get; set; } = new List<UserMenuDto>();
}

public class UserMenuDto
{
    public string? UserID { get; set; }
    public int MenuID { get; set; }
    public int ParentMenuID { get; set; }
    public string? Title { get; set; }
    public bool IsAssigned { get; set; }
    public List<UserMenuDto> SubMenu { get; set; } = new List<UserMenuDto>();
}

public class MenuItemViewModel
{
    public int MenuID { get; set; }
    public string? Title { get; set; }
    public IEnumerable<UserMenuDto>? SubMenu { get; set; } = new List<UserMenuDto>();
}