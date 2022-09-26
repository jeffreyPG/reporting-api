using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace reports.Services.DocumentStorage
{
    public interface IDocumentStorage
    {
        byte[] Get(string id);

        void Save(string id, byte[] bytes);

        void Delete(string id);

    }
}