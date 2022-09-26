using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlToOpenXml.Extensions;
using Shouldly;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class ListsTests
    {

        [TestMethod]
        public void GetOrderListLevelName()
        {
            var SUT = 0;
            var result = SUT.GetOrderedListType();
            result.ShouldBe("decimal");

            SUT = 1;
            result = SUT.GetOrderedListType();
            result.ShouldBe("lower-alpha");

            SUT = 2;
            result = SUT.GetOrderedListType();
            result.ShouldBe("lower-roman");

            SUT = 3;
            result = SUT.GetOrderedListType();
            result.ShouldBe("decimal");

        }

    }
}
