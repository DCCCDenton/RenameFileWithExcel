using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Data
{
    internal struct MergedCellBoundingBox
    {
        public CellSize MergedSize { get; set; }
        public int LeftColumn { get; set; }
        public int RightColumn { get; set; }
        public int TopRow { get; set; }
        public int BottomRow { get; set; }
    }
}
