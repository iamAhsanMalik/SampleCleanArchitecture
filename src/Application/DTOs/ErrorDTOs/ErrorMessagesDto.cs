namespace Application.DTOs.ErrorDTOs;

public class ErrorMessagesDto
{
    public List<ErrorMessageItem> ErrorMessage = new List<ErrorMessageItem>();
    public bool IsFailed;

    //public ErrorMessages();
}
public class ErrorMessageItem
{
    public bool IsMandatory;

    //public ErrorMessageItem();

    public string? DocType { get; set; }
    public string? Message { get; set; }
}