namespace Application.Extensions;
internal static class Common
{
    #region Linq Extensions
    // Foreach loop with index number support
    public static IEnumerable<(T element, int index)> WithIndex<T>(this IEnumerable<T> inputCollection) =>
        inputCollection.Select((item, index) => (item, index));
    #endregion
    #region Common Conversion Extension Methods
    public static string ToStringAiu(this object objectInput)
    {
        var result = objectInput != null ? Convert.ToString(objectInput) : default;
        return result ?? "";
    }
    public static bool ToBoolAiu(this string stringInput)
    {
        try
        {
            if (!string.IsNullOrEmpty(stringInput))
            {
                var result = stringInput.Trim().ToLower();
                if (!string.IsNullOrEmpty(result))
                {
                    return result is "-1" or "1" or "yes" or "true";
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
    }
    public static bool ToBoolAiu(this object objectInput)
    {
        try
        {
            if (objectInput != null)
            {
                var result = Convert.ToString(objectInput)?.Trim().ToLower();
                if (result != "")
                {
                    return result is "-1" or "1" or "yes" or "true";
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
    }
    public static bool IsDateAiu(this object objectInput)
    {
        try
        {
            if (objectInput != null)
            {
                var result = Convert.ToString(objectInput) ?? "";
                if (result != string.Empty)
                {
                    DateTime dateTime = DateTime.Parse(result);
                    return true;
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
    }
    public static bool IsNullDateAiu(this DateTime dateInput)
    {
        try
        {
            return dateInput.Year == 1900 && dateInput.Month == 1 && dateInput.Day == 1;
        }
        catch
        {
            return false;
        }
    }
    public static bool IsDbNullAiu(this object objectInput)
    {
        try
        {
            if (objectInput == null)
            {
                return true;
            }
            var result = Convert.ToString(objectInput) ?? "";
            return result == null || result.Length <= 0;
        }
        catch
        {
            return true;
        }
    }
    public static int ToIntAiu(this string stringInput)
    {
        try
        {
            if (!string.IsNullOrEmpty(stringInput))
            {
                return int.Parse(stringInput);
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    public static bool IsNullAiu(this object objectInput)
    {
        try
        {
            if (objectInput == null)
            {
                return true;
            }
            var result = Convert.ToString(objectInput) ?? "";
            return result?.Length == 0;
        }
        catch
        {
            return true;
        }
    }
    public static int ToIntAiu(this object integerInput)
    {
        try
        {
            if (integerInput != null)
            {
                var result = Convert.ToString(integerInput);
                int nm = result != null ? int.Parse(result) : 0;
                return nm;
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    public static DateTime ToDateTimeAiu(this string stringInput)
    {
        DateTime retDate = Convert.ToDateTime("1/1/1990 00:00:00");
        try
        {
            var result = Convert.ToString(stringInput);
            if (!string.IsNullOrEmpty(result))
            {
                retDate = DateTime.Parse(result);
            }
        }
        catch
        {
        }
        return retDate;
    }
    public static DateTime ToDateTimeAiu(this object objectInput)
    {
        DateTime retDate = Convert.ToDateTime("1/1/1990 00:00:00");
        try
        {
            if (objectInput != null)
            {
                var result = Convert.ToString(objectInput);
                if (result != string.Empty)
                {
                    retDate = result != null ? DateTime.Parse(result) : retDate;
                }
            }
        }
        catch
        {
        }
        return retDate;
    }
    public static decimal ToDecimalAiu(this string stringInput)
    {
        try
        {
            if (!string.IsNullOrEmpty(stringInput))
            {
                return decimal.Parse(stringInput);
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    public static decimal ToDecimalAiu(this object objectInput)
    {
        try
        {
            if (objectInput != null)
            {
                var result = Convert.ToString(objectInput);
                return result != null ? decimal.Parse(result) : 0;
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    public static double ToDoubleAiu(this string stringInput)
    {
        try
        {
            if (!string.IsNullOrEmpty(stringInput))
            {
                return double.Parse(stringInput);
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    public static double ToDoubleAiu(this object objectInput)
    {
        try
        {
            if (objectInput != null)
            {
                var result = Convert.ToString(objectInput);
                double nm = result != null ? double.Parse(result) : 0;
                return nm;
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }
    #endregion
}
