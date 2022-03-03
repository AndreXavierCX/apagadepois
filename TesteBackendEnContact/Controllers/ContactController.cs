using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TesteBackendEnContact.Controllers.Models;
using TesteBackendEnContact.Core.Domain.ContactBook;
using TesteBackendEnContact.Core.Interface.Contact;
using TesteBackendEnContact.Core.Interface.ContactBook;
using TesteBackendEnContact.Repository.Interface;

namespace TesteBackendEnContact.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ContactController : ControllerBase
    {
        private readonly ILogger<ContactController> _logger;

        public ContactController(ILogger<ContactController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> EnviarArquivo(IFormFile arquivos, [FromServices] IContactRepository contactRepository)
        {
            return Ok( await contactRepository.UploadCsvAsync(arquivos));
        }

        [HttpGet("{pesq}/{page}")]
        public async Task<IEnumerable<IContact>> SearchContactAsync(string pesq, int page, [FromServices] IContactRepository contactRepository)
        {
            return await contactRepository.SearchContractsAsync(pesq, page);
        }

        [HttpGet("{empresa}")]
        public async Task<IEnumerable<IContact>> PesquisarEmpresa(string empresa, [FromServices] IContactRepository contactRepository)
        {
            return await contactRepository.PesquisarEmpresa(empresa);
        }

        [HttpGet("exportar")]
        public async Task<IActionResult> ExportarContatos([FromServices] IContactRepository contactRepository)
        {
            return Ok(await contactRepository.ExportarContatos());
        }
    }
}
