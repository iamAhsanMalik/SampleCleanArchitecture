namespace Application.Contracts.Helpers;

public interface ICommonHelpers
{
    string ConvertToGrade(double thefinalmark);
    string ConvertToPoint(string tmp_grade, double tmp_credit_hrs);
    string ConvertToPointV2(string tmp_grade);
    string GetAccessLevelID(string _access_level);
    string GetAccessLevelName(string _access_level);
    string GetAssessmentTypeName(string strType);
    string GetStatusName(string strStatus);
    string GetStudentClassStatusName(int strStatus = 2);
    string GetSubmittedAssessmentStatusName(string strStatus);
}