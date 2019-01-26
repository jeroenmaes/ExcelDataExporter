using System.IO;

namespace ExcelToXmlParser.Services
{
    public interface IExcelConverter
    {
        string GetXML(Stream contents);
        string GetXML(Stream contents, string rootNodeName, string parentNodeName, string Xmlnamespace);
    }
}