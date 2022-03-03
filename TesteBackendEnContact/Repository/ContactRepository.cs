using CsvHelper;
using CsvHelper.Configuration;
using Dapper.Contrib.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TesteBackendEnContact.Core.Domain.Contact;
using TesteBackendEnContact.Core.Interface.Contact;
using TesteBackendEnContact.Database;
using TesteBackendEnContact.Repository.Interface;
using System.ComponentModel.DataAnnotations.Schema;
using TableAttribute = Dapper.Contrib.Extensions.TableAttribute;
using KeyAttribute = System.ComponentModel.DataAnnotations.KeyAttribute;
using Dapper;
using Excel = Microsoft.Office.Interop.Excel;

namespace TesteBackendEnContact.Repository
{
    public class ContactRepository : IContactRepository
    {

        private readonly DatabaseConfig databaseConfig;

        public ContactRepository(DatabaseConfig databaseConfig)
        {
            this.databaseConfig = databaseConfig;
        }

        public async Task<String> ExportarContatos()
        {
            using (var connection = new SqliteConnection(databaseConfig.ConnectionString))
            {
                var query = $"Select * from Contact; ";
                var result = await connection.QueryAsync<IContact>(query);

                connection.Close();

                // Inicia o componente Excel
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                //Cria uma planilha temporária na memória do computador
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //incluindo dados
                xlWorkSheet.Cells[1, 1] = "Id";
                xlWorkSheet.Cells[1, 2] = "ContactBookId";
                xlWorkSheet.Cells[1, 3] = "CompanyId";
                xlWorkSheet.Cells[1, 4] = "Name";
                xlWorkSheet.Cells[1, 5] = "Phone";
                xlWorkSheet.Cells[1, 6] = "Email";
                xlWorkSheet.Cells[1, 7] = "Address";

                int cont = 2;

                foreach (var item in result)
                {
                    xlWorkSheet.Cells[cont, 1] = item.Id;
                    xlWorkSheet.Cells[cont, 2] = item.ContactBookId;
                    xlWorkSheet.Cells[cont, 3] = item.CompanyId;
                    xlWorkSheet.Cells[cont, 4] = item.Name;
                    xlWorkSheet.Cells[cont, 5] = item.Phone;
                    xlWorkSheet.Cells[cont, 6] = item.Email;
                    xlWorkSheet.Cells[cont, 7] = item.Address;
                    cont++;
                }
                

                //Salva o arquivo de acordo com a documentação do Excel.
                xlWorkBook.SaveAs("arquivo.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                //o arquivo foi salvo na pasta Meus Documentos.
                string caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
               
                return "Concluído. Verifique em " + caminho + @"\dataexportContact.xls";
            }
        }

        public async Task<IEnumerable<IContact>> PesquisarEmpresa(string empresa)
        {
            using (var connection = new SqliteConnection(databaseConfig.ConnectionString))
            {
                var query = $"Select a.* from Contact a inner join Company b on a.CompanyId = b.Id Where b.Name = '" + empresa + "';";
                var result = await connection.QueryAsync<IContact>(query);

                connection.Close();
                return result;
            }
        }

        public async Task<IEnumerable<IContact>> SearchContractsAsync(string pesq, int page)
        {
            int pagesize = 2;

            using (var connection = new SqliteConnection(databaseConfig.ConnectionString))
            {
                var where = "Where a.Name like '%" + pesq + "%' or a.Phone like '%" + pesq + "%' or a.Email like '%" + pesq + "%' or a.Address like '%" + pesq + "%' or B.Name Like '%" + pesq + "%' or c.Name Like '%" + pesq + "%';";
                var query = $"Select DISTINCT a.* from Contact a inner join Company b on a.CompanyId = b.Id inner join ContactBook c on a.ContactBookId = b.Id " + where;
                var result = await connection.QueryAsync<IContact>(query);

                connection.Close();
                return result.Skip((page - 1) * pagesize).Take(pagesize);
            }
        }

        public async Task<object> UploadCsvAsync(IFormFile arquivos)
        {

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
            };
            if (arquivos == null)
            {
                return new { mensagem = "Erro" };
            }
            using (var reader = new StreamReader(arquivos.OpenReadStream()))
            using (var csv = new CsvReader(reader, config))
            {
                var arquivo = csv.GetRecords<ContactModel>().ToList();
                foreach (var arq in arquivo)
                {
                    using (var connection = new SqliteConnection(databaseConfig.ConnectionString))
                    {
                        ContactDao dao = new ContactDao(0, arq.ContactBookId, arq.CompanyId, arq.Name, arq.Phone, arq.Email, arq.Address);

                        try
                        {
                            dao.Id = await connection.InsertAsync(dao);
                        }
                        catch (Exception)
                        {
                            continue;
                        }                            

                    }
                }

            }

            string a = "sucesso";

            return new { a };

        }




        [Table("Contact")]
        public class ContactDao: IContact
        {
            [Key]
            public int Id { get; set; }
            public int ContactBookId { get; set; }
            public int CompanyId { get; set; }
            public string Name { get; set; }
            public string Phone { get; set; }
            public string Email { get; set; }
            public string Address { get; set; }

            public ContactDao(int id, int contactBookId, int companyId, string name, string phone, string email, string address)
            {
                Id = id;
                ContactBookId = contactBookId;
                CompanyId = companyId;
                Name = name;
                Phone = phone;
                Email = email;
                Address = address;
            }


            //int id, int contactbookid, int companyid, string name, string phone, string email, string address
            public IContact Export() => new Contact(Id, ContactBookId, CompanyId, Name, Phone, Email, Address);
        }

        public class ContactModel
        {
            public int Id { get; set; }
            [Required]
            public int ContactBookId { get; set; }
            [Required]
            public int CompanyId { get; set; }
            [Required]
            [StringLength(50)]
            public string Name { get; set; }
            [Required]
            [StringLength(20)]
            public string Phone { get; set; }
            [Required]
            [StringLength(50)]
            public string Email { get; set; }
            [Required]
            [StringLength(100)]
            public string Address { get; set; }
        }
    }
}
