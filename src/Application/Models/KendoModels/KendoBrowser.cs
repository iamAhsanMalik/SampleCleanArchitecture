namespace Application.Models.KendoModels;

/// <summary>
/// A modal that will be used for kendo Image and File Browser
/// </summary>
public class KendoBrowser
{
    public string? Name { get; set; }
    public long Size { get; set; }
    public string? EntryType { get; set; }
}
