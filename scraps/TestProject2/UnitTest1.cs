using ClosedXML.Excel;
using LodgeiT;
using Xunit.Abstractions;


namespace TestProject2
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
            Thread.Sleep(1300);
            _t.WriteLine("hello");
            Thread.Sleep(1300);
            _t.WriteLine("hello");
            Thread.Sleep(1300);
            _t.WriteLine("hello");
            Thread.Sleep(1300);
            _t.WriteLine("hello");
            _t.WriteLine("xUnit seems to be incapable of displaying console output while a test is running. Bummer.");
        }
    }
}
