namespace Application.DTOs;
public class ExchangeRateHistoryDto
{
    public string? CurrencyISOCode { get; set; }
    public DateTime FromDate { get; set; }
    public DateTime ToDate { get; set; }
    public string? Table { get; set; }
    public string? ErrMessage { get; set; }
}