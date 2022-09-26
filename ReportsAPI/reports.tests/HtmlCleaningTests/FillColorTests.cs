using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlToOpenXml.Extensions;
using Shouldly;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class FillColorTests
    {

        [TestMethod]
        public void SomeFillColorTests()
        {
            var SUT = "<p><line color=\"f4f6c6\" thickness=\"6\" /></p>";
            var result = SUT.FillColorOrDefault();
            result.ShouldBe("#f4f6c6");
        }

    }
}
