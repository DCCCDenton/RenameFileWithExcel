using ClosedXML.Excel;
using Microsoft.VisualBasic.Logging;
using RenameFileWithExcel.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Services
{
    internal class ExcelService
    {
        public XLWorkbook OpenExcelFile(string path)
        {
            using FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XLWorkbook workbook = new(stream);
            return workbook;
        }
        public List<ExcelCell> ReadExcel(XLWorkbook workbook, string workSheetName)
        {            
            IXLWorksheet worksheet = workbook.Worksheet(workSheetName);
            var excelBook = ReadAllUsedRangeFromSheetNew(worksheet);
            return excelBook;
        }
        public List<ExcelCell> ReadAllUsedRangeFromSheetNew(IXLWorksheet worksheet, int startRow = -1, int startCol = -1, int endRow = -1, int endCol = -1)
        {
            var usedRange = GetUsedAndMergedRangesOrNull(worksheet);
            var rows = usedRange.Rows();
            var result = new List<ExcelCell>();
            if (startRow == -1)
            {
                startRow = 0;
            }
            if (startCol == -1)
            {
                startCol = 0;
            }
            if (endRow == -1)
            {
                endRow = rows.Count();
            }
            if (endCol == -1)
            {
                endCol = rows.First().Cells().Count();
            }

            for (int i = startRow; i < endRow; i++)
            {
                var cells = rows.ElementAt(i).Cells();
                for (int j = startCol; j < endCol; j++)
                {
                    IXLColumn column = worksheet.Column(cells.ElementAt(j).Address.ColumnNumber);
                    IXLRow row1 = worksheet.Row(cells.ElementAt(j).Address.RowNumber);
                    if (column.IsHidden == false && row1.IsHidden == false)
                    {
                        string cellStringValue = null;
                        double? cellDoubleValue = null;
                        XLDataType dataType = cells.ElementAt(j).DataType;
                        ExcelCellTypeValue valueType;
                        if (dataType == XLDataType.Number)
                        {
                            try
                            {
                                cellDoubleValue = cells.ElementAt(j).Value.GetNumber();
                            }
                            catch
                            {
                                cellDoubleValue = cells.ElementAt(j).CachedValue.GetNumber();
                            }
                            valueType = ExcelCellTypeValue.Double;
                        }
                        else
                        {
                            valueType = ExcelCellTypeValue.String;
                            if (dataType != XLDataType.Error && dataType != XLDataType.Blank)
                            {
                                try
                                {
                                    cellStringValue = cells.ElementAt(j).Value.ToString();
                                }
                                catch
                                {
                                    cellStringValue = cells.ElementAt(j).CachedValue.ToString();
                                }
                            }
                        }
                        var lineStyle = GetLineStyle(cells.ElementAt(j).Style.Border);
                        var isMerged = IsMergedCell(cells.ElementAt(j));
                        var isMergedTopLeft = IsMergedTopLeft(cells.ElementAt(j));
                        var mergedCellBox = GetMergedCellBoundingBox(cells.ElementAt(j));
                        var excelCell = new ExcelCell
                        {
                            Alignment = cells.ElementAt(j).Style.Alignment,
                            RowNumber = cells.ElementAt(j).Address.RowNumber,
                            ColNumber = cells.ElementAt(j).Address.ColumnNumber,
                            ValueAsDouble = cellDoubleValue,
                            ValueAsString = cellStringValue,
                            ValueType = valueType,
                            IsMerged = isMerged,
                            IsMergedTopLeft = isMergedTopLeft,
                            MergedCellBox = mergedCellBox,
                            CellSize = new CellSize(cells.ElementAt(j).WorksheetColumn().Width, cells.ElementAt(j).WorksheetRow().Height),
                            NumberFormat = cells.ElementAt(j).Style.NumberFormat.Format,
                            BorderLineStyle = lineStyle,
                            WrapText = true //TODO: fix it
                        };
                        result.Add(excelCell);
                    }
                }
            }
            return result;
        }
        private IXLRange GetUsedAndMergedRangesOrNull(IXLWorksheet worksheet)
        {
            try
            {
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    return null;
                }

                var allRanges = new List<IXLRange> { usedRange };

                foreach (var mergedRange in worksheet.MergedRanges)
                {
                    if (!mergedRange.FirstCell().Value.IsBlank)
                    {
                        allRanges.Add(mergedRange);
                    }
                }

                int firstRow = allRanges.Min(r => r.FirstRow().RowNumber());
                int lastRow = allRanges.Max(r => r.LastRow().RowNumber());
                int firstColumn = allRanges.Min(r => r.FirstColumn().ColumnNumber());
                int lastColumn = allRanges.Max(r => r.LastColumn().ColumnNumber());

                return worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            }
            catch (Exception e)
            {
                return null;
            }
        }
        public LineStyle GetLineStyle(IXLBorder styleBorder)
        {
            return (styleBorder is null)
                ? default
                : new LineStyle
                {
                    LeftBorder = ConvertToCustomBorderStyle(styleBorder.LeftBorder),
                    RightBorder = ConvertToCustomBorderStyle(styleBorder.RightBorder),
                    TopBorder = ConvertToCustomBorderStyle(styleBorder.TopBorder),
                    BottomBorder = ConvertToCustomBorderStyle(styleBorder.BottomBorder)
                };
        }
        private CellBorderStyle ConvertToCustomBorderStyle(XLBorderStyleValues borderStyle)
        {
            if (Enum.TryParse(borderStyle.ToString(), out CellBorderStyle customBorderStyle))
            {
                return customBorderStyle;
            }
            return CellBorderStyle.None;
        }
        public bool IsMergedCell(IXLCell cell)
        {
            var mergedRange = cell?.MergedRange();
            return mergedRange is not null;
        }

        public bool IsMergedTopLeft(IXLCell cell)
        {
            if (IsMergedCell(cell))
            {
                var topLeftCell = cell?.MergedRange()?.FirstCell();
                var result = cell?.Address?.Equals(topLeftCell?.Address);
                return result ?? false;
            }

            return false;
        }
        private MergedCellBoundingBox GetMergedCellBoundingBox(IXLCell cell)
        {
            var mergedRange = cell.MergedRange();
            if (mergedRange != null)
            {
                // Calculate the total width and height of the merged range
                double totalWidth = 0;
                double totalHeight = 0;
                foreach (var row in mergedRange.Rows())
                {
                    double rowHeight = row.Cells().Max(c => c.WorksheetRow().Height);
                    totalHeight += rowHeight;
                }
                foreach (var column in mergedRange.Columns())
                {
                    totalWidth += column.WorksheetColumn().Width;
                }

                return new MergedCellBoundingBox
                {
                    MergedSize = new CellSize(totalWidth, totalHeight),
                    LeftColumn = mergedRange.RangeAddress.FirstAddress.ColumnNumber,
                    RightColumn = mergedRange.RangeAddress.LastAddress.ColumnNumber,
                    TopRow = mergedRange.RangeAddress.FirstAddress.RowNumber,
                    BottomRow = mergedRange.RangeAddress.LastAddress.RowNumber
                };
            }
            return default;
        }
    }
}
