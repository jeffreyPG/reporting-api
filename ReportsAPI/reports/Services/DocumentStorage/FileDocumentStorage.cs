using System.IO;

namespace reports.Services.DocumentStorage
{
    public class FileDocumentStorage : IDocumentStorage
    {
        private readonly string _directory;

        public FileDocumentStorage(string directory)
        {
            _directory = directory.TrimEnd('/');
        }

        public byte[] Get(string id)
        {
            var bytes = File.ReadAllBytes($"{_directory}/{id}");
            return bytes;
        }

        public void Save(string id, byte[] bytes)
        {
            Directory.CreateDirectory(_directory);
            File.WriteAllBytes($"{_directory}/{id}", bytes);
        }

        public void Delete(string id)
        {
            if(File.Exists(id))
            {
                File.Delete($"{_directory}/{id}");
            }            
        }

    }
}