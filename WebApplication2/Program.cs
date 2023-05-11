using Microsoft.AspNetCore.Mvc;
using LodgeiT;
using VDS.RDF.Query.Expressions.Functions.Sparql.String;

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

app.MapPost("/xlsx2rdf", ([FromBody] string input_fn) =>
{
    string output_fn = input_fn + ".rdf";
    /*t = new RdfTemplate();
    t.ExtractSheetGroupData()*/


    return "ook";
    //return { "error": error_msg};
    //return { "result": output_fn};
})
.WithName("xlsx2rdf")
.WithOpenApi();


app.Run();
