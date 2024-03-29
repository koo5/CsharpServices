using ClosedXML.Excel;
using LodgeiT;

namespace TestProject1;


public class Tests
{
    private TextWriter c;

    [SetUp]
    public void Setup()
    {
        c = TestContext.Progress;
    }

    [Test]
    public void Test1()
    {
        c.WriteLine("hello");
        Thread.Sleep(300);
        c.WriteLine("hello");
        Thread.Sleep(300);
        c.WriteLine("hello");
    }

    [Test]
    public void Test2()
    {
        c.WriteLine("let's go");
        string root = "ic_ui:investment_calculator_sheets";
        LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Noto Serif");
        /*
        foreach (var fontFamily in SixLabors.Fonts.SystemFonts.Collection.Families)
            Console.WriteLine(fontFamily.Name);
        */
        
        string? datapath = Environment.GetEnvironmentVariable("CSHARPSERVICES_DATADIR");
        if (datapath != null)
            Environment.CurrentDirectory = datapath;
        else
            Environment.CurrentDirectory =  Path.GetFullPath("../../../../data");
        
        string path = Directory.GetCurrentDirectory();
        c.WriteLine("The current directory is {0}", path);
        
        var wb = new XLWorkbook("empty IC template.xlsx");
        
        RdfTemplate.Tw = c;
        LodgeiT.RdfTemplate.Tw = c;
        RdfTemplate t = new RdfTemplate(wb, root);
        
        t.ExtractSheetGroupData();
        c.WriteLine(C.root.PrettyString());
        Assert.Pass();
    }
}