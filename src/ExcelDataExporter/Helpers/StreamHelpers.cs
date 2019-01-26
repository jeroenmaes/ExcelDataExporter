using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelToXmlParser.Helpers
{
    public static class StreamHelpers
    {
        public static Stream ConvertStringToStream(string contents)
        {
            // convert string to stream
            byte[] byteArray = Encoding.UTF8.GetBytes(contents);
            //byte[] byteArray = Encoding.ASCII.GetBytes(contents);

            MemoryStream stream = new MemoryStream(byteArray);

            return stream;
        }        
    }
}
