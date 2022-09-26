using Microsoft.VisualStudio.TestTools.UnitTesting;
using reports.Extensions;
using Shouldly;

namespace reports.tests.HtmlDecode
{

    [TestClass]
    public class HtmlDecodeTests
    {

        [TestMethod]
        public void RemoveHeader()
        {
            var SUT = "2_10_2021-16_35_46-EQE%20Style%20Template.docx";
            var result = SUT.Decode();
            result.ShouldBe("2_10_2021-16_35_46-EQE Style Template.docx");
        }

    }
}
