using VT.EmployeeImpExp.Models;

namespace VT.EmployeeImpExp.Repository.Interfaces
{
    public interface IEmpOrgRepository
    {
        bool ExportData(List<Organisation> organisations, List<Employee> employees);
        public List<Employee> GetEmployees(int pageNumber, int pageSize);
        public List<Organisation> GetOrganisations(int pageNumber, int pageSize);
    }
}
