using System.IO;

namespace ExcelToXmlParser.Services
{
    public interface IExcelConverter
    {
        void LoadStream(Stream contents);
        string GetJSON();
        string GetXML();
        string GetXML(string rootNodeName, string parentNodeName, string xmlNamespace);
    }
}