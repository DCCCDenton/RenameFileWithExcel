using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RenameFileWithExcel.Data;
using RenameFileWithExcel.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.ViewModel
{
    public partial class RenameFileViewModel : ObservableObject
    {
        private string FilePath { get; set; }
        private string FolderPath { get; set; }
        private XLWorkbook Workbook { get; set; }
        private List<ExcelCell> ExcelContent { get; set; } = new();
        private ExcelService ExcelService { get; set; } = new();

        [ObservableProperty]
        private List<IXLWorksheet> worksheets;
        [ObservableProperty]
        private IXLWorksheet selectedWorksheet;
        [ObservableProperty]
        private int nameColumn = 1;
        [ObservableProperty]
        private int bpmColumn = 4;

        [RelayCommand]
        private void OpenExcelFile()
        {
            FilePath = FileService.OpenFileDialog();
            if (!string.IsNullOrEmpty(FilePath))
            {
                Workbook = ExcelService.OpenExcelFile(FilePath);
                Worksheets = Workbook.Worksheets.ToList();
                SelectedWorksheet = Worksheets.First();
            }
        }

        [RelayCommand]
        private void SelectFolder()
        {
            FolderPath = FileService.GetFolderPath();
        }
        [RelayCommand]
        private void Run()
        {
            RenameService renameService = new();
            ExcelContent = ExcelService.ReadExcel(Workbook, SelectedWorksheet.Name);
            renameService.RenameFiles(FolderPath, ExcelContent, NameColumn, BpmColumn);
        }
    }
}
