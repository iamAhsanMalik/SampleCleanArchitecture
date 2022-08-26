namespace Application.DTOs;

public class PostingDateDto
{
    public bool IsPosted { get; set; }
    public DateTime MaxDate { get; set; }
    public string? UserMenuID { get; set; }
    public DateTime PostingDate { get; set; }
    public string? Message { get; set; }
    public string? PinCode { get; set; }

}