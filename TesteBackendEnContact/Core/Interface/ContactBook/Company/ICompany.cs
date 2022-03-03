namespace TesteBackendEnContact.Core.Interface.ContactBook.Company
{
    public interface ICompany
    {
        int Id { get; }
        int ContactBookId { get; }
        string Name { get; }
    }
}
