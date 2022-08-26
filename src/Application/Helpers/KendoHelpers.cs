using Application.Contracts.Infrastructure;
using Application.Models.KendoModels;
using Microsoft.AspNetCore.Mvc;

namespace Application.Helpers;

/// <summary>
/// Kendo Helper Class to browse files and folders from specific directory
/// </summary>
public class KendoHelpers
{
    private readonly IFileHelpers _fileHelpers;

    public KendoHelpers(IFileHelpers fileHelpers)
    {
        _fileHelpers = fileHelpers;
    }

    /// <summary>
    /// A method that will return all the files present in the provided path
    /// </summary>
    /// <param name="path"></param>
    /// <param name="filter"></param>
    /// <returns>Return Files Collection</returns>
    public IEnumerable<KendoBrowser>? GetFiles(string path, params string[] filters)
    {
        var directory = new DirectoryInfo(path);
        //var extensions = (filter ?? "*").Split(new string[] { ", ", ",", "; ", ";" }, StringSplitOptions.RemoveEmptyEntries);
        var extensions = filters?.Length > 0 ? filters : default;
        return extensions?.SelectMany(directory.GetFiles)
             .Select(file => new KendoBrowser
             {
                 Name = file.Name,
                 Size = file.Length,
                 EntryType = new KendoEntity().File
             });

    }
    /// <summary>
    /// A method that will return all the folders present in the provided path
    /// </summary>
    /// <param name="path"></param>
    /// <returns>Dirctory List</returns>
    public IEnumerable<KendoBrowser> GetDirectories(string path)
    {
        var directory = new DirectoryInfo(path);

        return directory.GetDirectories()
            .Select(subDirectory => new KendoBrowser
            {
                Name = subDirectory.Name,
                EntryType = new KendoEntity().Directory

            });
    }
    /// <summary>
    /// Kendo helper method that will browse files and folders in the provided path
    /// </summary>
    /// <returns>Files and Folders Array in JSON Form</returns>
    /// <exception cref="Exception"></exception>
    public JsonResult KendoFilesReader(string folderName, params string[] filters)
    {
        //var livePath = LiveSystemPath?.GetLivePath(folderName);
        var livePath = "";

        var fullPath = _fileHelpers.PathNormalizer(livePath);

        if (!Directory.Exists(fullPath))
        {
            Directory.CreateDirectory(fullPath);
        }

        if (_fileHelpers.IsAuthorizeRead(fullPath))
        {
            try
            {
                var files = GetFiles(fullPath, filters);
                var directories = GetDirectories(fullPath);
                var result = files?.Concat(directories);

                return new JsonResult(result!.ToArray());
            }
            catch (DirectoryNotFoundException)
            {
                throw new Exception("File Not Found");
            }
        }

        throw new Exception("Forbidden");
    }
    /// <summary>
    /// Kendo file uploader built for kendo editor specifically. For generic purpose, use FileUploaderAsync instead of this one
    /// </summary>
    /// <param name="file"></param>
    /// <returns>Kendo Browser Model in JSON Format</returns>
    /// <exception cref="Exception"></exception>
    public async Task<ActionResult> KendoFilesUploaderAsync(IFormFile file, string folderName)
    {

        //var livePath = LiveSystemPath.GetLivePath(folderName);
        var livePath = "";
        var fullPath = _fileHelpers.PathNormalizer(livePath);

        if (_fileHelpers.IsAuthorizeUpload(fullPath, file))
        {
            var fileName = await _fileHelpers.FileUploaderAsync(file, folderName);

            var result = new KendoBrowser
            {
                Size = file.Length,
                Name = fileName
            };

            return new JsonResult(result);
        }

        throw new Exception("Forbidden");
    }
    /// <summary>
    /// Kendo folder creator built for kendo editor specifically
    /// </summary>
    /// <param name="entry - KendoBrowser Model"></param>
    /// <param name="folderName"></param>
    /// <returns>Kendo Browser Model in JSON Format</returns>
    /// <exception cref="Exception"></exception>
    public JsonResult KendoDirectoryCreator(KendoBrowser entry, string folderName)
    {
        //var livePath = LiveSystemPath.GetLivePath(folderName);
        var livePath = "";

        var fullPath = _fileHelpers.PathNormalizer(livePath);
        string? name = entry.Name;

        if (name != null && _fileHelpers.IsAuthorizeCreateDirectory(fullPath, name))
        {
            var physicalPath = _fileHelpers.PathNormalizer(Path.Combine(fullPath, name));

            if (!Directory.Exists(physicalPath))
            {
                Directory.CreateDirectory(physicalPath);
            }

            return new JsonResult(entry);
        }

        throw new Exception("Forbidden");
    }
    /// <summary>
    /// Kendo file and folders remover
    /// </summary>
    /// <param name="entry - KendoBrowser Model"></param>
    /// <param name="folderName"></param>
    /// <returns>An empty JSON object</returns>
    /// <exception cref="Exception"></exception>
    public JsonResult KendoFilesRemover(KendoBrowser entry, string folderName)
    {
        //var livePath = LiveSystemPath.GetLivePath(folderName);
        var livePath = "";


        var fullPath = _fileHelpers.PathNormalizer(livePath);
        if (entry != null)
        {
            fullPath = Path.Combine(fullPath, entry.Name);

            if (entry.EntryType == new KendoEntity().File)
            {
                _fileHelpers.DeleteFile(fullPath);
            }
            else
            {
                _fileHelpers.DeleteDirectory(fullPath);
            }

            return new JsonResult(new object[0]);
        }

        throw new Exception("File Not Found");
    }
}
