using ClosedXML.Excel;
using LodgeiT;

namespace CsharpServices.Test
{
	public class UnitTest1
	{
		[Fact]
		public void Test1()
		{

			string root = "ic_ui:investment_calculator_sheets";

			LoadOptions.DefaultGraphicEngine = new ClosedXML.Graphics.DefaultGraphicEngine("Noto Serif");
			/*
			foreach (var fontFamily in SixLabors.Fonts.SystemFonts.Collection.Families)
				Console.WriteLine(fontFamily.Name);
			*/
			
			string? datapath = Environment.GetEnvironmentVariable("CSHARPSERVICES_DATADIR");
			if (datapath != null)
				Environment.CurrentDirectory = datapath;
			string path = Directory.GetCurrentDirectory();
			Console.WriteLine("The current directory is {0}", path);
			var wb = new XLWorkbook("empty IC template.xlsx");

			RdfTemplate t = new RdfTemplate(wb, root);
			if (!t.ExtractSheetGroupData(""))
			{
				throw new Exception(t.alerts);
				//Assert.False(result, "1 should not be prime");
			}
			string rdfStr = t.Serialize();            
		}
	}
}