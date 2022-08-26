namespace Application.DTOs.ErrorDTOs;

public class ErrorDTO
{
    public string? RequestId { get; set; }

    public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
}
