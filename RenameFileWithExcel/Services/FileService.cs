using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenameFileWithExcel.Services
{
    internal static class FileService
    {
        public static string GetFolderPath()
        {
            string folderPath = null;
            VistaFolderBrowserDialog dialog = new();
            dialog.RootFolder = Environment.SpecialFolder.MyComputer;
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog() == true)
            {
                folderPath = dialog.SelectedPath;
            }
            return folderPath;
        }

        public static string OpenFileDialog()
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "setting files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;
            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
            }
            return filePath;
        }
    }
}
