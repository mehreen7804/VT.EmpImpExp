namespace VT.EmployeeImpExp.Models
{
    public class Employee
    {
        public int Id{ get;  set; }
        public string? FirstName { get;  set; }
        public string? LastName { get;  set; }
        public int OrganisationId { get; set; }
        public string? OrganisationNumber { get; internal set; }
    }
}