using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ReportsAPISite.Services.ResourceProvider
{
    class EmbeddedResourceProvider : IResourceProvider
    {
        private readonly Assembly _assembly;

        public EmbeddedResourceProvider()
        {
            _assembly = Assembly.GetExecutingAssembly();
        }

        public List<Tuple<string, string>> GetAllFilesOfTypeInFolder(string folderPath, string fileType)
        {
            var scriptNames = _assembly
                    .GetManifestResourceNames()
                    .Where(r => r.StartsWith(folderPath) && r.EndsWith(fileType))
                    .OrderBy(x => x)
                    .ToList()
                ;

            var scriptsFound = new List<Tuple<string, string>>();
            foreach (var scriptName in scriptNames)
            {
                scriptsFound.Add(new Tuple<string, string>(scriptName, GetStringResource(scriptName)));
            }

            return scriptsFound;
        }

        public string GetStringResource(string resourceName)
        {
            var stream = _assembly.GetManifestResourceStream(resourceName);

            if (stream == null)
            {
                var message = $"embedded resource {resourceName} not found.  perhaps you forgot to mark it as an embedded resource?";
                throw new Exception(message);
            }

            var textStream = new StreamReader(stream);
            var output = textStream.ReadToEnd();
            return output;
        }
    }
}