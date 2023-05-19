#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Animation;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;


#endregion

namespace ArchSmarter_Addin_BonusExcelReader
{
    [Transaction(TransactionMode.Manual)]
    public class CmdInteropExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // Prompt user to select Excel file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel file|*.xls;*.xlsx;&.xlsm";
            selectFile.InitialDirectory = "C:\\Users\\TomH\\OneDrive - VEBH Architects\\Documents\\Revit Local";
            selectFile.Multiselect = false;

            // Create file variable
            string excelFile = "";
            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                excelFile = selectFile.FileName;
            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Select and Excel file");
                return Result.Failed;
            }

            ////// Open the Excel file
            ////Excel.Application excel = new Excel.Application();
            ////Excel.Workbook workbook = excel.Workbooks.Open(excelFile);

            ////// Get Furniture Set Data
            ////Excel.Worksheet worksheet = workbook.Worksheets[1];
            ////Excel.Range excelRange = (Excel.Range)worksheet.UsedRange;

            ////// Get row and column count
            ////int rows = excelRange.Rows.Count;
            ////int cols = excelRange.Columns.Count;

            ////// Read the data into a list
            ////List<List<String>> excelData = new List<List<String>>();
            ////for (int i = 1;i <= rows; i++)
            ////{
            ////    List<String> rowList = new List<String>();
            ////    for (int j = 1; j <= cols; j++)
            ////    {
            ////        string cellData = worksheet.Cells[i,j].Value.ToString();
            ////        rowList.Add(cellData);
            ////    }
            ////    excelData.Add(rowList);
            ////}

            ////// Close Excel
            ////excel.Quit();
            ///
                        // Open the Excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range excelRange = (Excel.Range)worksheet.UsedRange;

            // Get row and column count
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //Read the data into a list
            List<List<String>> excelData = new List<List<String>>();
            for (int i = 1; i <= rows; i++)
            {
                List<String> rowList = new List<String>();
                for (int j = 1; j <= cols; j++)
                {
                    string cellData = worksheet.Cells[i, j].Value.ToString();
                    rowList.Add(cellData);
                }
                excelData.Add(rowList);
            }

            //Create new worksheet
            Excel.Worksheet newWorkSheet = workbook.Worksheets.Add();
            newWorkSheet.Name = "Test Interop.Excel";

            //Write data to Excel
            for (int k = 1; k <= 10; k++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    newWorkSheet.Cells[j, k].Value = "Row " + j.ToString() + ": Column " + k.ToString();
                }
            }

            //Save and close Excel
            workbook.Save();
            excel.Quit();

            return Result.Succeeded;
        }        

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
