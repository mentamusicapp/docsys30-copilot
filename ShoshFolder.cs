using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    public class ShoshFolder:Folder
    {
        internal int shoshNum { get; set; }

        public ShoshFolder(int id, string description, string shortDescription, bool isMain, Branch b, bool isActive, FileType type, Classification classification, int shoshNum):
            base(id, description, shortDescription, isMain, b, isActive, type, classification)
        {
            this.shoshNum = shoshNum;
        }
    }
}
