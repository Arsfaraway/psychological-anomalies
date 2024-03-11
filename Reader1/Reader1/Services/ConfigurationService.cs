using Reader1.Interfaces;
using Reader1.Models.Configuration;
using Reader1.Storage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader1.Services
{
    public static class ConfigurationService
    {
        private static IStorage storage;
        static ConfigurationService() {
            storage = new FileStorage();
        }

        public static bool IsConfigured()
        {
            return IsConfigured(storage.GetConfig());
        }

        public static bool IsConfigured(Configuration config)
        {
            if (config == null) return false;
            return IsValid(config);
        }

        private static bool IsValid(Configuration config)
        {
            var properties = config.GetType().GetProperties();

            foreach (var property in properties)
            {
                var value = property.GetValue(config);

                if (value == null || (value is string && string.IsNullOrEmpty((string)value)))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
