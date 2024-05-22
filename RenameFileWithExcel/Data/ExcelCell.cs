using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Data
{
    internal class ExcelCell
    {
        public IXLAlignment Alignment { get; set; }
        public int RowNumber { get; set; }
        public int ColNumber { get; set; }
        public object Value { get; set; }

        public string ValueAsString { get; set; }
        public double? ValueAsDouble { get; set; }

        public CellSize CellSize { get; set; }

        public ExcelCellTypeValue ValueType { get; set; }
        public string NumberFormat { get; set; }
        public LineStyle BorderLineStyle { get; set; }
        public CellTextAlign TextAlign { get; set; }
        public bool WrapText { get; set; }
        public bool IsMerged { get; set; }
        public bool IsMergedTopLeft { get; set; }
        public MergedCellBoundingBox MergedCellBox { get; set; }
    }
}
