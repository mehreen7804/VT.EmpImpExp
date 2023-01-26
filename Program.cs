using OfficeOpenXml;
using VT.EmployeeImpExp.DataContext;
using VT.EmployeeImpExp.Repository;
using VT.EmployeeImpExp.Repository.Interfaces;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddSingleton<DapperContext>();
builder.Services.AddScoped<IEmpOrgRepository, EmpOrgRepository>();

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseAuthorization();

app.MapControllers();

app.Run();
