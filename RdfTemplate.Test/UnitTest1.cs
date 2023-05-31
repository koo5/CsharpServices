using ClosedXML.Excel;
using LodgeiT;
using Xunit.Abstractions;

namespace CsharpServices.Test
{
	public class UnitTest1
	{
	    private readonly ITestOutputHelper _t;
		public UnitTest1(ITestOutputHelper testOutputHelper)
		{
			_t = testOutputHelper;
		}
		[Fact]
		public void Test1()
		{
			_t.WriteLine("hello");
			Thread.Sleep(1000);
			_t.WriteLine("hello");
			Thread.Sleep(1000);
			_t.WriteLine("hello");
			Thread.Sleep(1000);
			_t.WriteLine("hello");
			Thread.Sleep(300);
			_t.WriteLine("hello");
			Thread.Sleep(300);
			_t.WriteLine("hello");
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
			_t.WriteLine("The current directory is {0}", path);
			
			var wb = new XLWorkbook("empty IC template.xlsx");

			// RdfTemplate t = new RdfTemplate(wb, root);
			// t._t = _t.WriteLine;
			// LodgeiT.C._t = _t.WriteLine;
			// if (!t.ExtractSheetGroupData(""))
			// {
			// 	throw new Exception(t.alerts);
			// 	//Assert.False(result, "1 should not be prime");
			// }
			// string rdfStr = t.Serialize();            
		}
	}
}