using System;
using System.IO;
using System.Text;

namespace ExcelToXmlParser.Services
{
    public class FolderFileStorage : IFileStorage
    {
        public Stream GetFileStream(string path)
        {
            var sr = new StreamReader(File.Open(path, FileMode.Open), Encoding.UTF8);
            return sr.BaseStream;
        }

        public void SaveFile(string path, Stream stream)
        {
            using (var fileStream = File.Create(path))
            {
                if (stream.CanSeek)
                    stream.Seek(0, SeekOrigin.Begin);

                stream.CopyTo(fileStream);
            }
        }
    }
}