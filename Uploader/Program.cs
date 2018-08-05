using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Uploader
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApplication.Workbooks.Add();
            var worksheet = (Worksheet)workbook.Worksheets.get_Item(1);

            #region Save
            
            worksheet.Cells[1, 1] = "Test";

            var saveFileDialog = new SaveFileDialog();
            string filename = "";
            saveFileDialog.Title = "Template file";
            saveFileDialog.Filter = "Excel Files (*.xslsx)|*.xlsx";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.AddExtension = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                filename = saveFileDialog.FileName;
            }

            workbook.SaveAs(filename);
            workbook.Close();
            excelApplication.Quit();
            #endregion

            var openFileDialog = new OpenFileDialog();
            string openFilename = "";
            openFileDialog.Title = "Template file";
            openFileDialog.Filter = "Excel Files (*.xslsx)|*.xlsx";
            openFileDialog.DefaultExt = "xlsx";
            openFileDialog.AddExtension = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                openFilename = openFileDialog.FileName;
            }

            var application = new Microsoft.Office.Interop.Excel.Application();
            var openedWorkBook = excelApplication.Workbooks.Open(openFilename);
            var openedWorksheet = (Worksheet)openedWorkBook.Worksheets.get_Item(1);

            var value = openedWorksheet.Cells[1][1].Value.ToString();
            Console.WriteLine(value);
            Console.ReadKey();
            openedWorkBook.Close();
            application.Quit();

        }
    }
}
