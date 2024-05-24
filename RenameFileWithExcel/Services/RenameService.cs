using RenameFileWithExcel.Data;
using RenameFileWithExcel.Services.Interface;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Services
{
    internal class RenameService : IRenameService
    {
        public void RenameFiles(string folderPath, List<ExcelCell> excelContent, int nameColumn, int bpmColumn)
        {
            string[] filesPath = Directory.GetFiles(folderPath, "*.mp3", SearchOption.AllDirectories);
            foreach (string filePath in filesPath)
            {
                try
                {
                    string prefix = FindBPM(excelContent, filePath, nameColumn, bpmColumn);
                    RenameFile(filePath, prefix);
                }
                catch { }                
            }
        }
        private void RenameFile(string filePath, string prefix = "",  string suffix = "")
        {
            FileInfo fileInfo = new(filePath);
            string oldFileName = fileInfo.Name;
            string newFileName = prefix + "-" + oldFileName;
            string newFilePath = fileInfo.DirectoryName + @"/" + newFileName;
            fileInfo.MoveTo(newFilePath);
        }

        private string FindBPM(List<ExcelCell> excelContent, string filePath, int nameColumn, int bpmColumn)
        {
            FileInfo fileInfo = new(filePath);
            string fileName = fileInfo.Name;
            int rowNumber = excelContent.Where(x => x.ColNumber == nameColumn).Where(x => x.ValueAsString == fileName).Select(x => x.RowNumber).First();
            string bpm = excelContent.Where(x => x.ColNumber == bpmColumn).First(x => x.RowNumber == rowNumber).ValueAsDouble.ToString();
            return bpm;
        }
    }
}
