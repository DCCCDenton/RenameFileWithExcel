using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Data
{
    internal struct CellTextAlign
    {
        public CellTextAlign(TextAlignHorizon h, TextAlignVertical v)
        {
            Horizontal = h;
            Vertical = v;
        }

        public TextAlignHorizon Horizontal { get; set; }
        public TextAlignVertical Vertical { get; set; }
    }
}
