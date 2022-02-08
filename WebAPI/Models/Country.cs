using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI.Models
{
    public class Country
    {
        public Name name { get; set; }
        public List<string> capital { get; set; }
        public string region { get; set; }
        public string subregion { get; set; }
        public double area { get; set; }
        public int population { get; set; }
        public List<string> timezones { get; set; }
        public Flags flags { get; set; }
    }
}