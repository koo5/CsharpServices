using Microsoft.AspNetCore.Mvc;
using LodgeiT;
using VDS.RDF.Query.Expressions.Functions.Sparql.String;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.MapPost("/xlsx2rdf", ([FromBody] string input_fn, string root) =>
{
    string output_fn = input_fn + ".rdf";
    RdfTemplate t = new RdfTemplate(new XLWorkbook(input_fn), root);
    if (!t.ExtractSheetGroupData(""))
        return { "error:":t.alerts};
    t.SerializeToFile(output_fn);  
    return { "result": "ok"};
})
.WithName("xlsx2rdf")
.WithOpenApi();


app.Run();
