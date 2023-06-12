using Microsoft.AspNetCore.Mvc;
using LodgeiT;
using ClosedXML.Excel;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Logging.SetMinimumLevel(LogLevel.Trace);
builder.Services.AddLogging();
builder.Logging.AddConsole();
builder.Logging.AddDebug();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}


app.Logger.LogError("ERROR!");
app.Logger.LogWarning("WARN!");
app.Logger.LogInformation("INFO!");
app.Logger.LogDebug("DEBUG!");
app.Logger.LogTrace("TRACE!");

RdfTemplate.Tw = Console.Out;


app.MapGet("/health", () => "ok")
    .WithName("health")
    .WithOpenApi();


//app.MapPost("/xlsx_to_rdf", ([FromBody] string root, [FromBody] string input_fn, [FromBody] string output_fn) =>
app.MapPost("/xlsx_to_rdf", ([FromBody] RpcRequest rrr) =>
    {
        app.Logger.LogInformation("INFO!");
        LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Noto Serif");
        
        RdfTemplate t = new RdfTemplate(new XLWorkbook(rrr.input_fn), rrr.root);
        
        if (!t.ExtractSheetGroupData(""))
            return new RpcReply (null, t.Alerts );
        t.SerializeToFile(rrr.output_fn);
        
        // refactor: this is a hack to reset the trace variables
        C.root = null;
        C.current_context = null;
        
        return new RpcReply ("ok",null );
    
    })
    .WithName("xlsx_to_rdf")
    .WithOpenApi()
    .WithDescription("called by frontend. returns either result or error. error is a text with newlines, possibly including a rendering of a user-centric backtrace given by t.alerts");



app.Run("http://0.0.0.0:17789");




internal record RpcRequest(string root, string input_fn, string output_fn)
{
}
internal record RpcReply(string? result, string? error)
{
}

