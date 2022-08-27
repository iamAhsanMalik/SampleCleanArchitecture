using Application.Contracts.Helpers;

namespace Application.Helpers;
internal class CommonHelpers : ICommonHelpers
{
    private int RoundedDistance(int value, double dividedBy)
    {
        return (int)decimal.Round(Convert.ToDecimal(value / dividedBy), MidpointRounding.AwayFromZero);
    }

    private bool InRange(int low, int high, int value)
    {
        return value >= low && value <= high;
    }

    public string GetAssessmentTypeName(string strType)
    {
        string retVal = "";
        if (strType == "1")
        {
            retVal = "Assignment";
        }
        else if (strType == "2")
        {
            retVal = "Essay";
        }
        else if (strType == "3")
        {
            retVal = "Quiz";
        }
        else if (strType == "4")
        {
            retVal = "Worksheet";
        }
        else if (strType == "5")
        {
            retVal = "Survey";
        }
        else
        {
            retVal = strType;
        }
        return retVal;
    }
    public string GetStatusName(string strStatus)
    {
        string retVal = "";
        if (strStatus == "0")
        {
            retVal = "New";
        }
        else if (strStatus == "1")
        {
            retVal = "Active";
        }
        else if (strStatus == "2")
        {
            retVal = "Inactive";
        }
        else
        {
            retVal = strStatus;
        }
        return retVal;
    }

    public string GetSubmittedAssessmentStatusName(string strStatus)
    {
        string retVal = "";
        if (strStatus == "0")
        {
            retVal = "Ungraded";
        }
        else if (strStatus == "1")
        {
            retVal = "Graded";
        }
        else if (strStatus == "2")
        {
            retVal = "Incomplete";
        }
        else if (strStatus == "3")
        {
            retVal = "Redo";
        }
        else
        {
            retVal = strStatus;
        }
        return retVal;
    }

    public string GetStudentClassStatusName(int strStatus = 2)
    {
        if (strStatus == 1)
        {
            return "Active";
        }
        else if (strStatus == 2)
        {
            return "Inactive";
        }
        else if (strStatus == 3)
        {
            return "Withdrawn";
        }
        else if (strStatus == 4)
        {
            return "On-hold";
        }
        else if (strStatus == 5)
        {
            return "Completed";
        }
        else if (strStatus == 6)
        {
            return "Locked";
        }
        else
        {
            return "Inactive";
        }
    }

    public string GetAccessLevelName(string _access_level)
    {
        string retVal = "";
        if (_access_level == "2")
        {
            retVal = "Students";
        }
        else if (_access_level == "3")
        {
            retVal = "Parents";
        }
        else if (_access_level == "5")
        {
            retVal = "Affiliates";
        }
        else if (_access_level == "10")
        {
            retVal = "Reps";
        }
        else if (_access_level == "20")
        {
            retVal = "Teachers";
        }
        else if (_access_level == "50")
        {
            retVal = "Principals";
        }
        else if (_access_level == "75")
        {
            retVal = "Companies";
        }
        else if (_access_level == "99")
        {
            retVal = "Administrators";
        }
        return retVal;
    }

    public string GetAccessLevelID(string _access_level)
    {
        string retVal = "";
        if (_access_level == "Students")
        {
            retVal = "2";
        }
        else if (_access_level == "Parents")
        {
            retVal = "3";
        }
        else if (_access_level == "Affiliates")
        {
            retVal = "5";
        }
        else if (_access_level == "Reps")
        {
            retVal = "10";
        }
        else if (_access_level == "Teachers")
        {
            retVal = "20";
        }
        else if (_access_level == "MiniAdmins")
        {
            retVal = "50";
        }
        else if (_access_level == "Administrators")
        {
            retVal = "99";
        }
        return retVal;
    }
    public string ConvertToGrade(double thefinalmark)
    {
        string finalgrade = "";
        try
        {
            if (thefinalmark >= 89.5 && thefinalmark <= 100)
            {
                finalgrade = "A";
            }
            else if (thefinalmark >= 79.5 && thefinalmark <= 89.4)
            {
                finalgrade = "B";
            }
            else if (thefinalmark >= 69.5 && thefinalmark <= 79.4)
            {
                finalgrade = "C";
            }
            else if (thefinalmark >= 59.5 && thefinalmark <= 69.4)
            {
                finalgrade = "D";
            }
            else if (thefinalmark < 59.4)
            {
                finalgrade = "F";
            }
        }
        catch (Exception)
        {
            finalgrade = "";
        }
        return finalgrade;
    }

    public string ConvertToPoint(string tmp_grade, double tmp_credit_hrs)
    {

        double tmp_point = 0;
        if (tmp_grade == "A")
        {
            if (tmp_credit_hrs < 1)
            {
                if (tmp_credit_hrs >= .1)
                {
                    tmp_point = 2;
                }
                else
                {
                    tmp_point = 0;
                }
            }
            else
            {
                tmp_point = 4;
            }
        }
        else if (tmp_grade == "B")
        {
            if (tmp_credit_hrs < 1)
            {
                tmp_point = 1.5;
            }
            else
            {
                tmp_point = 3;
            }
        }
        else if (tmp_grade == "C")
        {
            if (tmp_credit_hrs < 1)
            {
                tmp_point = 1;
            }
            else
            {
                tmp_point = 2;
            }
        }
        else if (tmp_grade == "D")
        {
            if (tmp_credit_hrs < 1)
            {
                tmp_point = .5;
            }
            else
            {
                tmp_point = 1;
            }
        }
        else if (tmp_grade == "F")
        {
            tmp_point = 0;
        }
        return tmp_point.ToString();
    }

    public string ConvertToPointV2(string tmp_grade)
    {
        double tmp_point = 0;
        if (tmp_grade == "A")
        {
            tmp_point = 4;
        }
        else if (tmp_grade == "B")
        {
            tmp_point = 3;
        }
        else if (tmp_grade == "C")
        {
            tmp_point = 2;
        }
        else if (tmp_grade == "D")
        {
            tmp_point = 1;
        }
        else if (tmp_grade == "F")
        {
            tmp_point = 0;
        }
        return tmp_point.ToString();
    }
    /// <summary>
    /// Converts a byte array to a string of hex characters
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    protected static string GetString(byte[] data)
    {
        StringBuilder results = new StringBuilder();

        foreach (byte b in data)
        {
            results.Append(b.ToString("X2"));
        }

        return results.ToString();
    }

    /// <summary>
    /// Converts a string of hex characters to a byte array
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    protected static byte[] GetBytes(string data)
    {
        // GetString() encodes the hex-numbers with two digits
        byte[] results = new byte[data.Length / 2];

        for (int i = 0; i < data.Length; i += 2)
        {
            results[i / 2] = Convert.ToByte(data.Substring(i, 2), 16);
        }

        return results;
    }
}
