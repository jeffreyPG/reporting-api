using HtmlToOpenXml.Services.HtmlCleaning;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace reports.tests.HtmlCleaningTests
{
    [TestClass]
    public class HeaderRemovalTests
    {
        [TestMethod]
        public void RemoveHeader()
        {
            var html = "<html><body><!-- <header>test</header> --></body></html>";
            var SUT = new HtmlCleanerBuilder(html)
                            .RemoveHeaderContent()
                            .Build();
            var result = "<html><body></body></html>";
            SUT.ShouldBe(result);
        }

    }
}
