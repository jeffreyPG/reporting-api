using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace reports.Models
{
    /// <summary>
    /// 
    /// </summary>
    public class ChartsModel
    {
        /// <summary>
        /// 
        /// </summary>
        public string ViewId { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public int Sort { get; set; }
    }

    public class AuthRequest
    {
        public CredetialsBody credentials { get; set; }
    }

    public class CredetialsBody
    {
        public string name { get; set; }

        public string password { get; set; }

        public Dictionary<string, string> site { get; set; }
    }
}