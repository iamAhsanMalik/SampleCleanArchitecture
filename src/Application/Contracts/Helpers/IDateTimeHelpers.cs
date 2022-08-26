namespace Application.Contracts.Helpers;

public interface IDateTimeHelpers
{
    string ConvertDecimalToHoursMinutes(double time);
    void GetFirstWeekOfYear(int year, out DateTime weekStart, out DateTime weekStop);
    string GetMonthName(int month, bool abbreviate = false, IFormatProvider? provider = null);
    void GetWeekInterval(DateTime date, out DateTime dtWeekStart, out DateTime dtWeekStop);
    void GetWeekInterval(int year, int weekNo, out DateTime weekStart, out DateTime weekStop);
    List<KeyValuePair<DateTime, DateTime>>? GetWeekSeries(DateTime toDate);
    List<KeyValuePair<DateTime, DateTime>>? GetWeekSeries(DateTime fromDate, DateTime toDate);
    DateTime GetWeekStart(DateTime date);
    DateTime GetWeekStop(DateTime date);
    string PluraliseTime(int num, string word);
}