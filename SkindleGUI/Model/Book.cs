using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SkindleGUI.Model
{
    class Book
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string Name { get; set; }

        public override string ToString()
        {
            if (Name != null) return Name;
            return FileName;
        }
    }
}
