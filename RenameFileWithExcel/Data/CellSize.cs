using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Data
{
    internal struct CellSize
    {
        public CellSize(double width, double height)
        {
            WidthInLogicalPixel = width;
            HeightInLogicalPixel = height;
        }

        public double WidthInLogicalPixel { get; set; }
        public double HeightInLogicalPixel { get; set; }
    }
}
