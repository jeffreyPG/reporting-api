using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlToOpenXml.Extensions;
using Shouldly;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class QuillJsStringTests
    {

        [TestMethod]
        public void GetLevelIndentFrom()
        {
            var SUT = "ql-indent-1";
            var result = SUT.GetQuillJSIndentLevel();
            result.ShouldBe(1);

            SUT = "ql-indent-94";
            result = SUT.GetQuillJSIndentLevel();
            result.ShouldBe(94);
        }

    }
}
