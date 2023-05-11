using ClosedXML.Excel;
using LodgeiT;

namespace CsharpServices.Test
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {

            var wb = new XLWorkbook("C:/test.xlsx");
            var t = new RdfTemplate(wb, );
            t.ExtractSheetGroupData()

            t.test11();
            //Assert.False(result, "1 should not be prime");
        }
    }
}