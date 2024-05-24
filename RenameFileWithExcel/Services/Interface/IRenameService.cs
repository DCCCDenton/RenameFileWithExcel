using RenameFileWithExcel.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Services.Interface
{
    internal interface IRenameService
    {
        public void RenameFiles(string folderPath, List<ExcelCell> excelContent, int nameColumn, int bpmColumn);
    }
}
