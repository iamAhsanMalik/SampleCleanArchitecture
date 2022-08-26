namespace Application.DTOs;
public class ImportDepositFileDto
{
    public int Id { get; set; }
    public bool IsError { get; set; }
    public bool IsFileImported { get; set; }
    public string? ErrorMessage { get; set; }
    public bool IsRowValid { get; set; }
    public string? RowErrMessage { get; set; }
    public bool IsPosted { get; set; }
    public string? CpUserRefNo { get; set; }
    public string? Sheet { get; set; }
    public string? ErrMessage { get; set; }
    public string? AgentPrefix { get; set; }
    public int CustomerId { get; set; }
    public string? CustomerName { get; set; }
    public string? Currency { get; set; }
    public decimal Amount { get; set; }
    public decimal Rate { get; set; }
    public string? Type { get; set; }
    public string? Details { get; set; }
    public string? Account { get; set; }
    public decimal FCAmount { get; set; }
    public decimal LCAmount { get; set; }
    public string? ExcelFilePath { get; set; }
    public int RowNumber { get; set; }
    public string? Name { get; set; }
    public string? Extension { get; set; }
    public long Size { get; set; }
    public DateTime PostingDate { get; set; }
    public List<ImportDepositFileDto> UploadFilesList { get; set; } = new List<ImportDepositFileDto>();

}

