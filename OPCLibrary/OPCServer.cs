using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OPCLibrary
{
    public class OPCServer
    {
        public OPCServer()
        {
            Name = "default";
        }

        ~OPCServer()
        {

        }

        public string Name
        { get; set; }
    }
}
