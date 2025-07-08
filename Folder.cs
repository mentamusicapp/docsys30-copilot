using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    public class Folder
    {
        public int id { get; internal set; }
        public string description { get; internal set; }
        internal string shortDescription { get; set; }
        internal bool isMain { get; set; }
        internal Branch branch { get; set; }
        internal bool isActive { get; set; }
        internal FileType type { get; set; }
        internal Classification classification { get; set; }

        /// Initializes a new instance of the Folder class.
         
        public Folder(int id, string description, string shortDescription, bool isMain, Branch branch, bool isActive, FileType type, Classification classification)
        {
            this.id = id;
            this.description = description;
            this.shortDescription = shortDescription;
            this.isMain = isMain;
            this.branch = branch;
            this.isActive = isActive;
            this.type = type;
            this.classification = classification;
        }
    }
}
