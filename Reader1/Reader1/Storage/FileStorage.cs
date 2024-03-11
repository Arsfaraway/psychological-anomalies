using Reader1.Interfaces;
using Reader1.Models.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;


namespace Reader1.Storage
{
    public class FileStorage : IStorage
    {
        private string fileName = "config.txt";
        public FileStorage() { }

        public Configuration GetConfig()
        {
            if (!File.Exists(fileName)) { return null; }
            return JsonSerializer.Deserialize<Configuration>(File.ReadAllText(GetFilePath()));
        }

        public void SaveConfig(Configuration config)
        {

            File.WriteAllText(GetFilePath(), JsonSerializer.Serialize(config));
        }

        private string GetFilePath()
        {
            return Path.Combine(System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), fileName);
        }
    }


}
