namespace Application.Contracts.Infrastructure;
public interface IFileHelpers
{
    bool IsAuthorizeDeleteFile(string path);
    bool CanAccessPath(string path);
    void DeleteDirectory(string path);
    void DeleteFile(string path);
    Task<string> EmailTemplatesReaderAsync(string fileNameWithExtension);
    Task<string> FileUploaderAsync(IFormFile file, string folderName);
    string[] ImageFilters(params string[] inputFilters);
    bool IsAuthorizeCreateDirectory(string path, string name);
    bool IsAuthorizeDeleteDirectory(string path);
    bool IsAuthorizeRead(string path);
    bool IsAuthorizeUpload(string path, IFormFile file);
    bool IsValidImage(string fileName);
    string PathNormalizer(string path);
}