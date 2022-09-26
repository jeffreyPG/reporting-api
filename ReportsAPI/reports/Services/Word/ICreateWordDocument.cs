using reports.Endpoints.CreateWordDocumentController.Models;

namespace reports.Services.Word
{
    public interface ICreateWordDocument
    {
        string CreateWordDocument(CreateArgs args, string fileName);

        string ReplaceDocumentStyles(CreateArgs args, string fileName);

    }
}