using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaInspector.Logic
{
    public class ClassInfoDTO
    {
        public string Namespace { get; set; }
        public string ClassName { get; set; }
        public List<string> Fields { get; set; }
        public List<string> Properties { get; set; }
        public List<MethodInfoDto> Methods { get; set; }
        public List<string> Delegates { get; set; }
        public List<string> Events { get; set; }
        public List<string> Enums { get; set; } 
    }
}
