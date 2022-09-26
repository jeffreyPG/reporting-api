namespace ReportsAPISite.Services.DocumentStorage
{
    public interface IDocumentStorage
    {
        byte[] Get(string id);
        void Save(string id, byte[] bytes);
        void Delete(string id);
    }
}