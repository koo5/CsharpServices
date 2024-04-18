using Microsoft.AspNetCore.Mvc;
using LodgeiT;
using ClosedXML.Excel;

using DocumentFormat.OpenXml;  
using DocumentFormat.OpenXml.Packaging;  
using DocumentFormat.OpenXml.Spreadsheet;  


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


app.Logger.LogError(		"test error from Program.cs");
app.Logger.LogWarning(		"test warn from Program.cs");
app.Logger.LogInformation(	"test info from Program.cs");
app.Logger.LogDebug(		"test debug from Program.cs");
app.Logger.LogTrace(		"test trace from Program.cs");

RdfTemplate.Tw = Console.Out;


app.MapGet("/health", () => "ok")
    .WithName("health")
    .WithOpenApi();


//app.MapPost("/xlsx_to_rdf", ([FromBody] string root, [FromBody] string input_fn, [FromBody] string output_fn) =>
app.MapPost("/xlsx_to_rdf", ([FromBody] xlsx_to_rdfRpcRequest rrr) =>
    {
        app.Logger.LogInformation("xlsx_to_rdf:");
        //LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Noto Serif");
        
        var openSettings = new OpenSettings()
        {
            RelationshipErrorHandlerFactory = package =>
            {
                return new UriRelationshipErrorHandler();
            }
        };
        
        File.Copy(rrr.input_fn, rrr.input_fn+"-fixup.xlsx", true);
        
        using (var doc = SpreadsheetDocument.Open(rrr.input_fn+"-fixup.xlsx", true, openSettings))
        {
            doc.Save();
        }       
       
        var w = new XLWorkbook(rrr.input_fn+"-fixup.xlsx");
		
		if (w == null)
		{
			return new RpcReply (null, "failed to resave file");
		}
        
        RdfTemplate t = null;      
		t = new RdfTemplate(w, new Uri(rrr.root));	
        
        if (!t.ExtractSheetGroupData(""))
            return new RpcReply (null, t.Alerts );
        t.SerializeToFile(rrr.output_fn);
        app.Logger.LogInformation("wrote " + rrr.output_fn);
                
        return new RpcReply ("ok",null );
    
    })
    .WithName("xlsx_to_rdf")
    .WithOpenApi()
    .WithDescription("called by frontend. returns either result or error. error is a text with newlines, possibly including a rendering of a user-centric backtrace given by t.alerts");



app.MapPost("/rdf_to_xlsx", ([FromBody] rdf_to_xlsxRpcRequest rrr) =>
    {
        app.Logger.LogInformation("rdf_to_xlsx:");

        var w = new XLWorkbook();
        RdfTemplate t = new RdfTemplate(w);
        t.LoadResultSheets(new StreamReader(File.OpenRead(rrr.input_file)));
        w.SaveAs(rrr.output_directory + "/result.xlsx");
        
        return new RpcReply ("ok", null);

    })
    .WithName("rdf_to_xlsx")
    .WithOpenApi();




app.Run("http://0.0.0.0:17789");




internal record xlsx_to_rdfRpcRequest(string root, string input_fn, string output_fn){}
internal record rdf_to_xlsxRpcRequest(string input_file, string output_directory){}
internal record RpcReply(string? result, string? error){}



public class UriRelationshipErrorHandler : RelationshipErrorHandler
{
    public override string Rewrite(Uri partUri, string id, string uri)
    {
        return "http://link-invalid.example.com";
    }
}

