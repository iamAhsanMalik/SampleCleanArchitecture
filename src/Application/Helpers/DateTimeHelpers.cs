using Application.Contracts.Helpers;

namespace Application.Helpers;
internal class DateTimeHelpers : IDateTimeHelpers
{
    public string GetMonthName(int month, bool abbreviate = false, IFormatProvider? provider = default)
    {
        var info = DateTimeFormatInfo.GetInstance(provider);
        return abbreviate ? info.GetAbbreviatedMonthName(month) : info.GetMonthName(month);
    }
    public string ConvertDecimalToHoursMinutes(double time)
    {
        TimeSpan timespan = TimeSpan.FromHours(time);
        int hours = timespan.Hours;
        int mins = timespan.Minutes;

        // Convert to hours and minutes
        string hourString = hours > 0 ? string.Format("{0} " + PluraliseTime(hours, "hour"), hours) : "";
        string minString = mins > 0 ? string.Format("{0} " + PluraliseTime(mins, "minute"), mins) : "";

        // Add a space between the hours and minutes if necessary
        return hours > 0 && mins > 0 ? hourString + " " + minString : hourString + minString;
    }

    /// <summary>Pluralise hour or minutes based on the amount of time</summary>
    /// <param name="num">The number of hours or minutes</param>
    /// <param name="word">The word to pluralise e.g. "hour" or "minute"</param>
    /// <returns> Returns correct English pluralisation e.g. 3 hours, 1 minute, 0 minutes</returns>
    public string PluraliseTime(int num, string word)
    {
        return num == 0 || num > 1 ? word + "s" : word;
    }
    public DateTime GetWeekStart(DateTime date)
    {
        DateTime weekStart;
        int monday = 1;
        int crtDay = (int)date.DayOfWeek;
        if (date.DayOfWeek == DayOfWeek.Sunday)
        {
            crtDay = 7;
        }

        int difference = crtDay - monday;
        weekStart = date.AddDays(-difference);
        return weekStart;
    }
    public DateTime GetWeekStop(DateTime date)
    {
        DateTime weekStart;
        int sunday = 7;
        int crtDay = (int)date.DayOfWeek;
        if (date.DayOfWeek == DayOfWeek.Sunday)
        {
            crtDay = 7;
        }

        int difference = sunday - crtDay;
        weekStart = date.AddDays(difference);
        return weekStart;
    }
    public void GetWeekInterval(int year, int weekNo,
        out DateTime weekStart, out DateTime weekStop)
    {
        GetFirstWeekOfYear(year, out weekStart, out weekStop);
        if (weekNo == 1)
        {
            return;
        }

        weekNo--;
        int daysToAdd = weekNo * 7;
        DateTime dt = weekStart.AddDays(daysToAdd);
        GetWeekInterval(dt, out weekStart, out weekStop);
    }
    public List<KeyValuePair<DateTime, DateTime>>? GetWeekSeries(DateTime toDate)
    {
        //gets week series from beginning of the year
        DateTime dtStartYear = new DateTime(toDate.Year, 1, 1);
        List<KeyValuePair<DateTime, DateTime>>? list = GetWeekSeries(dtStartYear, toDate);
        if (list?.Count > 0)
        {
            KeyValuePair<DateTime, DateTime> week = list[0];
            list[0] = new KeyValuePair<DateTime, DateTime>(dtStartYear, week.Value);
        }
        return list;
    }
    public List<KeyValuePair<DateTime, DateTime>>? GetWeekSeries(DateTime fromDate, DateTime toDate)
    {
        if (fromDate > toDate)
        {
            return default;
        }

        List<KeyValuePair<DateTime, DateTime>> list = new List<KeyValuePair<DateTime, DateTime>>(100);
        DateTime weekStart, weekStop;
        toDate = GetWeekStop(toDate);
        while (fromDate <= toDate)
        {
            GetWeekInterval(fromDate, out weekStart, out weekStop);
            list.Add(new KeyValuePair<DateTime, DateTime>(weekStart, weekStop));
            fromDate = fromDate.AddDays(7);
        }
        return list;
    }
    public void GetFirstWeekOfYear(int year, out DateTime weekStart, out DateTime weekStop)
    {
        DateTime date = new DateTime(year, 1, 1);
        GetWeekInterval(date, out weekStart, out weekStop);
    }
    public void GetWeekInterval(DateTime date,
        out DateTime dtWeekStart, out DateTime dtWeekStop)
    {
        dtWeekStart = GetWeekStart(date);
        dtWeekStop = GetWeekStop(date);
    }
}
