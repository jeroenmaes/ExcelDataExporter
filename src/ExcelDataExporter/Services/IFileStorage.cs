using System.IO;

namespace ExcelToXmlParser.Services
{
    public interface IFileStorage
    {
        void SaveFile(string path, Stream stream);

        Stream GetFileStream(string path);
    }
}