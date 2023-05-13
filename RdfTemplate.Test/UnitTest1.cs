using ClosedXML.Excel;
using LodgeiT;

namespace CsharpServices.Test
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {

            string root = "";
            var wb = new XLWorkbook("C:/test.xlsx");
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