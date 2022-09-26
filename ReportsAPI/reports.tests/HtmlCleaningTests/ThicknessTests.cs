using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlToOpenXml.Extensions;
using Shouldly;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class ThicknessTests
    {

        [TestMethod]
        public void SomeTicknessTests()
        {
            var SUT = "<p><line color=\"000000\" thickness=\"6\" /></p>";
            var result = SUT.ThicknessOrDefault();
            result.ShouldBe("6pt");
        }

    }
}
