namespace Application.DTOs;

public class UnAuthorizedVouchersDto
{
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? ErrMessage { get; set; }
    public string? Table { get; set; }
    public string? UnAuthorizedAllVouchers { get; set; }
    public string? UnAuthorizedCashVouchers { get; set; }
    public string? UnAuthorizedTransferVouchers { get; set; }
}