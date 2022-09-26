using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace reports.Models
{
    public class ProjectData
    {
        public string type { get; set; }
        public string username { get; set; }
        public string title { get; set; }
        public string orientation { get; set; }
        public Report[] report { get; set; }
    }

    public class Report
    {
        public string sheetName { get; set; }
        public Group[] data { get; set; }
    }

    public class Group
    {
        public string group { get; set; }
        public Section[] sections { get; set; }
    }

    public class Section
    {
        public string title { get; set; }
        public string[] content { get; set; }
    }
}