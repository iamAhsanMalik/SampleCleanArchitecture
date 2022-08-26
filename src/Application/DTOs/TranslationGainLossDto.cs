namespace Application.DTOs;

public class TranslationGainLossDto
{
    public string? Table { get; set; }
    public string? PinCode { get; set; }
    public string? CpUserRefNo { get; set; }
    public string? Memo { get; set; }
    public string? ErrMessage { get; set; }
    public bool IsPosted { get; set; }
    public string? CpRefID { get; set; }
    public bool HaveData { get; set; }
}