using Domain.Constants;

namespace Infrastructure.Files;
/// <summary>
/// File Helpers methods used to work with files such as FileUploading, Path Normalization etc..
/// </summary>
internal class FileHelpers : IFileHelpers
{
    private readonly IWebHostEnvironment _webHostEnvironment;
    private readonly IGeneratorHelpers _generatorHelpers;

    public FileHelpers(IWebHostEnvironment webHostEnvironment, IGeneratorHelpers generatorHelpers)
    {
        _webHostEnvironment = webHostEnvironment;
        _generatorHelpers = generatorHelpers;
    }
    public async Task<string> EmailTemplatesReaderAsync(string fileNameWithExtension) => await File.ReadAllTextAsync(PathNormalizer($"{_webHostEnvironment.ContentRootPath}/EmailTemplates/Auth/{fileNameWithExtension}"));


    /// <summary>
    /// Async file upload function which will save the file in the directory that you provide.
    /// </summary>
    /// <param name="file"></param>
    /// <param name="folderName: (Optional)"></param>
    /// <returns>FileName</returns>
    public async Task<string> FileUploaderAsync(IFormFile file, string folderName)
    {
        var directoryPath = LiveSystemPath.GetLivePath(folderName);
        directoryPath = PathNormalizer(directoryPath);
        try
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to create the folder. {ex.Message}");
        }
        var fileName = _generatorHelpers.GenerateRandomCode(5) + "_" + file.FileName;

        var fileSavePath = Path.Combine(directoryPath, fileName);

        fileSavePath = PathNormalizer(fileSavePath);
        try
        {
            using (var stream = new FileStream(fileSavePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to upload the file. {ex.Message}");
        }
        return fileName;
    }
    /// <summary>
    /// Path Normalizer which will trim and add double (//) slashes to the path
    /// </summary>
    /// <param name="Path"></param>
    /// <returns>Normalize Path (String)</returns>
    public string PathNormalizer(string path)
    {
        return Path.GetFullPath(new Uri(path).LocalPath)
                   .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    /// <summary>
    /// A method that can be used to check whether user / system is authrize to read files and foldes in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <returns>Boolean - True or False</returns>
    public bool IsAuthorizeRead(string path)
    {
        return CanAccessPath(path);
    }

    /// <summary>
    /// A method that can be used to check whether user / system is authrize to create folder in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <param name="name"></param>
    /// <returns>Boolean - True or False</returns>
    public bool IsAuthorizeCreateDirectory(string path, string name)
    {
        return CanAccessPath(path);
    }
    /// <summary>
    /// A method that can be used to check whether user / system is authrize to access specific path
    /// </summary>
    /// <param name="path"></param>
    /// <returns>Boolean - True or False</returns>
    public bool CanAccessPath(string path)
    {
        var rootPath = Path.GetFullPath(Path.Combine(path));

        return path.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase);
    }
    /// <summary>
    /// A method that can be used to delete specific file in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <exception cref="Exception"></exception>
    public void DeleteFile(string path)
    {
        if (!IsAuthorizeDeleteFile(path))
        {
            throw new Exception("Forbidden");
        }

        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
    /// <summary>
    /// A method that can be used to delete specific folder in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <exception cref="Exception"></exception>
    public void DeleteDirectory(string path)
    {
        if (!IsAuthorizeDeleteDirectory(path))
        {
            throw new Exception("Forbidden");
        }

        if (Directory.Exists(path))
        {
            Directory.Delete(path, true);
        }
    }
    /// <summary>
    /// A method that can be used to check whether user / system is authrize to delete file in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    public bool IsAuthorizeDeleteFile(string path)
    {
        return CanAccessPath(path);
    }
    /// <summary>
    /// A method that can be used to check whether user / system is authrize to delete folder in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <returns>Boolean - True or False</returns>
    public bool IsAuthorizeDeleteDirectory(string path)
    {
        return CanAccessPath(path);
    }
    /// <summary>
    /// A method that can be used to check whether user / system is authrize to upload file in a given path
    /// </summary>
    /// <param name="path"></param>
    /// <param name="file"></param>
    /// <returns>Boolean - True or False</returns>
    public bool IsAuthorizeUpload(string path, IFormFile file)
    {
        return CanAccessPath(path) && IsValidImage(file.FileName);

    }
    /// <summary>
    /// File validator that help use to valid files that can be uploaded
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns>Boolean - True or False</returns>
    public bool IsValidImage(string fileName)
    {
        var extension = Path.GetExtension(fileName);
        var allowedExtensions = ImageFilters(ImageType.PNG, ImageType.JPG, ImageType.JPEG, ImageType.Gif);

        return allowedExtensions.Any(e => e.Equals("*.*") || e.EndsWith(extension, StringComparison.OrdinalIgnoreCase));
    }
    /// <summary>
    /// An helper method for fileuploader that could be used to restrict type of files users could upload. For example, if you pass "png" as input to this method then user could only upload .png file. Use ImageType class to input parameters
    /// </summary>
    /// <param name="inputFilters - such as png, jpg, jpeg, gif"></param>
    /// <returns>Filter as string</returns>
    public string[] ImageFilters(params string[] inputFilters) => inputFilters;
    public static class LiveSystemPath
    {
        public static string GetLivePath(string folderName)
        {
            return $"/{folderName}/";
            //return $"C:/Users/Dlls/TestingUploads/{folderName}/";
        }
    }
}
