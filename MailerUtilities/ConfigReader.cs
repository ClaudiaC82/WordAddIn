using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace MailerUtilities
{
    public class ConfigReader
    {
        [Serializable]
        public class Configuration : PersistableObject
        {
            public string connection_string { get; set; }
            public string excel_path { get; set; }
            public string test2 { get; set; }
            public string test3 { get; set; }
        }

        public class PersistableObject
        {
            public static T Load<T>(string fileName) where T : PersistableObject, new()
            {
                T result = default(T);

                using (FileStream stream = File.OpenRead(fileName))
                {
                    result = new XmlSerializer(typeof(T)).Deserialize(stream) as T;
                }

                return result;
            }
        }
    }
}
