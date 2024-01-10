using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Xml.Linq;

namespace wordEditor.Models
{
    public class CustomImageInfo
    {
        public byte[] ImageBytes { get; set; }
        public string ContentType { get; set; }
        public string AltText { get; set; }

    }

}