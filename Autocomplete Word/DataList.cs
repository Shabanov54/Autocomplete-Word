using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Autocomplete_Word
{
    internal class DataList
    {
        [Key]
        internal string id { get; set; }
        internal string name { get; set; }
        internal string login { get; set; }
        internal string pass { get; set; }
        internal string dateTime { get; set; }
        internal string flag { get; set; }
    }
}
