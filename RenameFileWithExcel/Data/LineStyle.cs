using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Data
{
    internal struct LineStyle
    {
        public CellBorderStyle LeftBorder { get; set; }
        public CellBorderStyle RightBorder { get; set; }
        public CellBorderStyle TopBorder { get; set; }
        public CellBorderStyle BottomBorder { get; set; }
    }
}
