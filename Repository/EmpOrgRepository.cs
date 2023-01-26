using static VT.EmployeeImpExp.Repository.EmpOrgRepository;
using VT.EmployeeImpExp.DataContext;
using VT.EmployeeImpExp.Repository.Interfaces;
using VT.EmployeeImpExp.Models;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Drawing.Printing;
using Dapper;
using static ExcelImportAPI.Controllers.ExcelImpExportController;

namespace VT.EmployeeImpExp.Repository
{
    public class EmpOrgRepository : IEmpOrgRepository
    {
        private readonly DapperContext _context;

        public EmpOrgRepository(DapperContext context)
        {
            _context = context;
        }
        public List<Organisation> GetOrganisations(int pageNumber, int pageSize)
        {
            try
            {
                var query = "SELECT * FROM Organisation ORDER BY Id OFFSET @pageSize * (@pageNumber - 1) ROWS FETCH NEXT @pageSize ROWS ONLY";
                using (var db = _context.CreateConnection())
                {
                    return db.Query<Organisation>(query, new
                    {
                        pageNumber,
                        pageSize
                    }).ToList().ToList();

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Employee> GetEmployees(int pageNumber, int pageSize)
        {
            try
            {
                var query = "SELECT * FROM Employee ORDER BY Id OFFSET @pageSize * (@pageNumber - 1) ROWS FETCH NEXT @pageSize ROWS ONLY";
                using (var db = _context.CreateConnection())
                {
                    return db.Query<Employee>(query, new
                    {
                        pageNumber,
                        pageSize
                    }).ToList().ToList();
                    
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public bool ExportData(List<Organisation> organisations, List<Employee> employees)
        {
            bool isExported = false;
            try
            {
                using (var db = _context.CreateConnection())
                {
                    //(@"CREATE TABLE Organisation (Id INT IDENTITY(1,1) PRIMARY KEY, OrganisationName VARCHAR(255), OrganisationNumber VARCHAR(255), AddressLine1 VARCHAR(255), AddressLine2 VARCHAR(255), AddressLine3 VARCHAR(255), AddressLine4 VARCHAR(255), Town VARCHAR(255), Postcode VARCHAR(255), Number VARCHAR(255))");
                    //(@"CREATE TABLE Employee (Id INT IDENTITY(1,1) PRIMARY KEY, OrganisationId INT, FirstName VARCHAR(255), LastName VARCHAR(255), FOREIGN KEY (OrganisationId) REFERENCES Organisation(Id))");
                    db.Execute("INSERT INTO Organisation (OrganisationName, OrganisationNumber, AddressLine1, AddressLine2, AddressLine3, AddressLine4, Town, Postcode, Number) VALUES (@OrganisationName, @OrganisationNumber, @AddressLine1, @AddressLine2, @AddressLine3, @AddressLine4, @Town, @Postcode, @Number)", organisations);
                    var orgNumbers = organisations.Select(x => x.OrganisationNumber).ToList();
                    var orgIds = db.Query<OrgIdNum>("SELECT Id,OrganisationNumber FROM Organisation WHERE OrganisationNumber IN @orgNumbers", new { orgNumbers }).ToList();
                    for (int i = 0; i < employees.Count; i++)
                    {
                        var org = orgIds.Find(o => o.OrganisationNumber == employees[i].OrganisationNumber);
                        if (org != null)
                            employees[i].OrganisationId = org.Id;
                    }
                    db.Execute("INSERT INTO Employee (OrganisationId, FirstName, LastName) VALUES (@OrganisationId, @FirstName, @LastName)", employees);
                    isExported = true;
                }
            }
            catch (Exception ex)
            {
                isExported = false;
            }
            return isExported;
        }

    }
}
