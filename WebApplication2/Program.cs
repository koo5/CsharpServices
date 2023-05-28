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



/* called by frontend. returns either result or error. error is a text with newlines, possibly including a rendering
 of a user-centric backtrace given by t.alerts */
app.MapPost("/xlsx2rdf", (string root, /*[FromBody] */string input_fn) =>
{
    string output_fn = input_fn + ".rdf";
    RdfTemplate t = new RdfTemplate(new XLWorkbook(input_fn), root);
    if (!t.ExtractSheetGroupData(""))
        return new RpcReply (null, t.alerts );
    t.SerializeToFile(output_fn);
    return new RpcReply ("ok",null );
    
})
.WithName("xlsx2rdf")
.WithOpenApi();


app.MapGet("/weatherforecast", () =>
{
    var forecast = Enumerable.Range(1, 5).Select(index =>
        new WeatherForecast
        (
            DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
            Random.Shared.Next(-20, 55),
            "555"
        ))
        .ToArray();
    return forecast;
})
.WithName("GetWeatherForecast")
.WithOpenApi();



app.Run();



internal record WeatherForecast(DateOnly Date, int TemperatureC, string? Summary)
{
    public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);
}
internal record RpcReply(string? result, string? error)
{
}

