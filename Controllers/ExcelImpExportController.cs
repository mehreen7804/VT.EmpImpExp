using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using Dapper;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Microsoft.Extensions.Logging;
using VT.EmployeeImpExp.Models;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Data;
using VT.EmployeeImpExp.Repository.Interfaces;
using System.Drawing.Printing;
using System.Collections.Generic;

namespace ExcelImportAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelImpExportController : ControllerBase
    {
        private readonly ILogger<ExcelImpExportController> _logger;
        private readonly IEmpOrgRepository _empRepo;

        public ExcelImpExportController(ILogger<ExcelImpExportController> logger, IEmpOrgRepository empRepo)
        {
            _logger = logger;
            _empRepo = empRepo;
        }

        [HttpPost("ImportXlsxData/{fileName}")]
        public IActionResult ImportXlsxData(string fileName)
        {
            try
            {
                if (!System.IO.File.Exists(fileName))
                    return BadRequest("File not found");

                var organisations = new List<Organisation>();
                var employees = new List<Employee>();

                using (var package = new ExcelPackage(new FileInfo(fileName)))
                {
                    var orgSheet = package.Workbook.Worksheets["Organisation"];
                    var orgStartRow = orgSheet.Dimension.Start.Row;
                    var orgEndRow = orgSheet.Dimension.End.Row;

                    for (int row = orgStartRow + 1; row <= orgEndRow; row++)
                    {
                        organisations.Add(new Organisation
                        {
                            OrganisationName = orgSheet.Cells[row, 1].Value?.ToString(),
                            OrganisationNumber = orgSheet.Cells[row, 2].Value?.ToString(),
                            AddressLine1 = orgSheet.Cells[row, 3].Value?.ToString(),
                            AddressLine2 = orgSheet.Cells[row, 4].Value?.ToString(),
                            AddressLine3 = orgSheet.Cells[row, 5].Value?.ToString(),
                            AddressLine4 = orgSheet.Cells[row, 6].Value?.ToString(),
                            Town = orgSheet.Cells[row, 7].Value?.ToString(),
                            Postcode = orgSheet.Cells[row, 8].Value?.ToString(),
                            Number = orgSheet.Cells[row, 9].Value?.ToString()
                        });
                    }

                    var empSheet = package.Workbook.Worksheets["Employee"];
                    var empStartRow = empSheet.Dimension.Start.Row;
                    var empEndRow = empSheet.Dimension.End.Row;

                    for (int row = empStartRow + 1; row <= empEndRow; row++)
                    {
                        employees.Add(new Employee
                        {
                            OrganisationNumber = empSheet.Cells[row, 1].Value?.ToString(),
                            FirstName = empSheet.Cells[row, 2].Value?.ToString(),
                            LastName = empSheet.Cells[row, 3].Value?.ToString()
                        });
                    }
                }

                SaveXlsxData(organisations, employees);
                return Ok("Data imported and saved successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while importing data from Excel file");
                return StatusCode(500, "An error occurred while importing data from Excel file");
            }
        }
        public class OrgIdNum
        {
            public int Id { get; set; }
            public string OrganisationNumber { get; set; }
        }
        private bool SaveXlsxData(List<Organisation> organisations, List<Employee> employees)
        {
            return _empRepo.ExportData(organisations, employees);
        }

        [HttpGet("GetOrganisations/{pageNumber}/{pageSize}")]
        public List<Organisation> GetOrganisations(int pageNumber, int pageSize)
        {
            try
            {
                return _empRepo.GetOrganisations(pageNumber, pageSize);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while retrieving employees from database");
                return null;
            }
        }
        [HttpGet("GetEmployees/{pageNumber}/{pageSize}")]
        public List<Employee> GetEmployees(int pageNumber, int pageSize)
        {
            try
            {
               return _empRepo.GetEmployees(pageNumber, pageSize);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while retrieving employees from database");
                return null;
            }
        }
    }
}


