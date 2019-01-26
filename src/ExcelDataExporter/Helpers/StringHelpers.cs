using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExcelToXmlParser.Helpers
{
    public static class StringHelpers
    {      
        public static string CleanInvalidChars(this string text)
        {
            text = new string(text.Where(c => char.IsLetterOrDigit(c) || char.IsWhiteSpace(c) || c == '-').ToArray());
            
            return text;
        }

        public static string RemoveLineBreaks(this string text)
        {
            string replaceWith = "";
            string removedBreaks = text.Replace("\r\n", replaceWith).Replace("\n", replaceWith).Replace("\r", replaceWith);

            return text;
        }

        public static string CleanInvalidXmlChars(this string text)
        {
            text = text.Replace(" ", "_");
            text = text.Replace("(", "");
            text = text.Replace(")", "");

            return text;
        }

        public static string TrimWhiteSpaces(this string text)
        {
            text = text.TrimStart(' ');
            text = text.TrimEnd(' ');
            text = text.Trim();

            return text;
        }

        public static string ReverseDateString(this string text)
        {
            DateTime date;
            var result = DateTime.TryParse(text, out date);

            if (!result)
                return string.Empty;

            return date.ToString("yyyyMMdd");
        }
    }
}
