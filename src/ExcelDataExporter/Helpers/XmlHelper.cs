using System.IO;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelToXmlParser.Helpers
{
    public class XmlHelper
    {
        public static string Serialize(object obj)
        {
            using (var memoryStream = new MemoryStream())
            using (var reader = new StreamReader(memoryStream))
            {
                var serializer = new XmlSerializer(obj.GetType());
                serializer.Serialize(memoryStream, obj);
                memoryStream.Position = 0;
                return reader.ReadToEnd();
            }
        }

        public static string Serialize<T>(object obj)
        {
            using (var memoryStream = new MemoryStream())
            using (var reader = new StreamReader(memoryStream))
            {
                var serializer = new XmlSerializer(typeof(T));
                serializer.Serialize(memoryStream, obj);
                memoryStream.Position = 0;
                return reader.ReadToEnd();
            }
        }

        public static T Deserialize<T>(Stream stream)
        {
            var ser = new XmlSerializer(typeof(T));
            var result = (T)ser.Deserialize(stream);
            return result;
        }

        public static T Deserialize<T>(string xmlString)
        {
            using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
            {
                var ser = new XmlSerializer(typeof(T));
                var result = (T)ser.Deserialize(reader);

                return result;
            }            
        }

        public static MemoryStream SerializeToXmlStream<T>(object o)
        {
            var serializer = new XmlSerializer(typeof(T));

            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);

            writer.Write(serializer);
            writer.Flush();
            stream.Position = 0;

            return stream;
        }
    }
}
