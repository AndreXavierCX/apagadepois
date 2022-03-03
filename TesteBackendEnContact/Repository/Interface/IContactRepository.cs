using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TesteBackendEnContact.Core.Interface.Contact;

namespace TesteBackendEnContact.Repository.Interface
{
    public interface IContactRepository
    {
        Task<Object> UploadCsvAsync(IFormFile arquivos);

        Task<IEnumerable<IContact>> SearchContractsAsync(string pesq, int page);

        Task<IEnumerable<IContact>> PesquisarEmpresa(string pesq);

        Task<string> ExportarContatos();
    }
}
    