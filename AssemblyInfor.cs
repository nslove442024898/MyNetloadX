using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyNetloadX
{
    public class AssemblyInfor
    {
        public string CodeBase { get; set; }
        public string FullName { get; set; }
        public string Location { get; set; }
        public string FullyQualifiedName { get; set; }
        public string ScopeName { get; set; }
        public string ImageRuntimeVersion { get; set; }

        public AssemblyInfor(Assembly ass)
        {
            this.CodeBase = ass.CodeBase;
            this.FullName = ass.FullName;
            this.Location = ass.Location;
            this.FullyQualifiedName = ass.ManifestModule.FullyQualifiedName;
            this.ScopeName = ass.ManifestModule.ScopeName;
            this.ImageRuntimeVersion = ass.ImageRuntimeVersion;
        }
    }
}
