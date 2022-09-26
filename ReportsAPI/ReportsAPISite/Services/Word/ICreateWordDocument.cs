using ReportsAPISite.Endpoints.Word;

namespace ReportsAPISite.Services.Word
{
    public interface ICreateWordDocument
    {

        string CreateWordDocument(CreateArgs args, string fileName);
        string ReplaceDocumentStyles(CreateArgs args, string fileName);

    }
}