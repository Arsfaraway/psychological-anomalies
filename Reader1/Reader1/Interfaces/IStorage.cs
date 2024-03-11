using Reader1.Models.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader1.Interfaces
{
    public interface IStorage
    {
        Configuration GetConfig();
        void SaveConfig(Configuration config);
    }
}
