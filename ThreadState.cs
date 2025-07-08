using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    class ThreadState
    {
        public byte [] FileData { get; set; }
        public int FileName { get; set; }
        public string FileExt { get; set; }
        public string FileVers { get; set; }
    }
}
