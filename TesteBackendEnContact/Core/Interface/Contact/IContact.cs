using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TesteBackendEnContact.Core.Interface.Contact
{
    public class IContact
    {
        public int Id { get; }
        public int ContactBookId { get; }
        public int CompanyId { get; }
        public string Name { get; }
        public string Phone { get;}
        public string Email { get; }
        public string Address { get; }
    }
}
