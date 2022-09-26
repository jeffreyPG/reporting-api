using System;
using System.Collections.Generic;

namespace ReportsAPISite.Services.ResourceProvider
{
    public interface IResourceProvider
    {
        List<Tuple<string, string>> GetAllFilesOfTypeInFolder(string folderPath, string fileType);
        string GetStringResource(string resourceName);
    }
}
