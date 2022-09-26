using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlToOpenXml.Services.HtmlCleaning;
using Shouldly;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class LineBreakTests
    {

        [TestMethod]
        public void LineBreakTest()
        {
            var html = "<p><br></p>";
            var SUT = new HtmlCleanerBuilder(html)
                            .RemoveLineBreakInsideParagraphTag()
                            .Build();
            var result = "<p> </p>";
            SUT.ShouldBe(result);

        }

        [TestMethod]
        public void DualLineBreakTest()
        {
            var html = "<html><body><p><p>ABC</p><p><br></p><p>ABCD</p><p><br></p><p>ABCEF</p></p><p><p>dkaosndaskdnksdn</p><p>asdjkasbdkBAD</p><p><br></p><p>kasdksajdkbjaD</p></p></body></html>";
            var SUT = new HtmlCleanerBuilder(html)
                            .RemoveLineBreakInsideParagraphTag()
                            .Build();
            var result = "<html><body><p><p>ABC</p><p> </p><p>ABCD</p><p> </p><p>ABCEF</p></p><p><p>dkaosndaskdnksdn</p><p>asdjkasbdkBAD</p><p> </p><p>kasdksajdkbjaD</p></p></body></html>";
            SUT.ShouldBe(result);

        }

        [TestMethod]
        public void MultipleLineBreakTest()
        {
            var html = "<p></p><br>some test<br><br>some test<br><p><br><br></p><br> some test <br>more text blablabla <br> some test <br><p><br><br></p>more text blablabla";
            var SUT = new HtmlCleanerBuilder(html)
                            .RemoveLineBreakInsideParagraphTag()
                            .Build();
            var result = "<p></p><br>some test<br><br>some test<br><p><br></p><br> some test <br>more text blablabla <br> some test <br><p><br></p>more text blablabla";
            SUT.ShouldBe(result);

        }
    }
}
