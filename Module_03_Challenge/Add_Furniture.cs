#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using static Module_03_Challenge.Add_Furniture;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace Module_03_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Add_Furniture : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Prompt user to select Excel file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel file|*.xls;*.xlsx;&.xlsm";
            selectFile.InitialDirectory = "C:\\Users\\TomH\\OneDrive - VEBH Architects\\Documents\\Revit Local\\Revit Addins Project\\RAB_Module_03_Challenge_Files";
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


            // Open the Excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);


            // Get furniture Sets
            getExcelFileData fsList = new getExcelFileData(workbook, 1);
            List<List<string>> fsData = fsList.excelData;
            int fsCounter = fsData.Count;

            // Get furniture Types
            getExcelFileData ftList = new getExcelFileData(workbook, 2);
            List<List<string>> ftData = ftList.excelData;
            int ftCounter = ftData.Count;

            // Close Excel
            excel.Quit();


            // Get rooms
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert furniture");

                // Create a list of rooms
                foreach (SpatialElement room in collector)
                {
                    LocationPoint locPt = room.Location as LocationPoint;
                    XYZ roomPt = locPt.Point as XYZ;

                    //// Read furniture set parameter
                    string fSet = Utils.GetParameterValueAsString(room, "Furniture Set");

                    // Set Included furniture variable
                    string curIncFurn = "";

                    // Get the included furniture list
                    FurnitureSet furnitureSet = new FurnitureSet(fsCounter, fsData, fSet);
                    if (furnitureSet.match != "")
                    {
                        // Get the current furniture set and count
                        curIncFurn = furnitureSet.innerList[2];

                        FurnitureType furniture = new FurnitureType(doc, curIncFurn, ftData, ftCounter, roomPt);

                        // Update room furniture count
                        Utils.SetParameterValueAsDouble(room, "Furniture Count", furniture.furnCount);
                    }

                }
                t.Commit();
            }

            return Result.Succeeded;
        }


        public class getExcelFileData
        {
            public List<List<string>> excelData = new List<List<string>>();
            public getExcelFileData(Workbook curWb, int curWs)
            {
                Excel.Worksheet worksheet = curWb.Worksheets[curWs];
                Excel.Range excelRange = (Excel.Range)worksheet.UsedRange;

                // Get row and column count
                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                //Read the data into a list
                for (int i = 1; i <= rows; i++)
                {
                    List<string> rowData = new List<string>();
                    for (int j = 1; j <= cols; j++)
                    {
                        string cellContent = worksheet.Cells[i, j].Value.ToString();
                        rowData.Add(cellContent);
                    }
                    excelData.Add(rowData);
                }
            }
        }

        public class FurnitureSet
        {
            public int outerIndex = 1;
            public int innerIndex = 0;
            public int innerCounter = 1;
            public string curFS = "";
            public string match = "";
            public string curIncFurn = "";
            public List<string> innerList = new List<string>();

            public FurnitureSet(int _fsCounter, List<List<string>> _fsData, string _fSet)
            {
                for (int i = 1; i < _fsCounter; i++)
                {
                    innerList = _fsData[outerIndex];
                    curFS = innerList[0];
                    if (curFS == _fSet)
                    {
                        match = _fSet;
                        break;
                    }
                    else
                    {
                        match = "";
                        outerIndex++;
                    }
                }
            }
        }

        public class FurnitureType
        {
            public string curList { get; set; }
            public string[] splitString = new string[] { };
            public double furnCount = 0;

            public FurnitureType(Document doc, string _curList, List<List<string>> _ftData, int _ftCounter, XYZ _roomPt)
            {
                curList = _curList;
                splitString = curList.Split(',');
                for (int j = 0; j < splitString.Length; j++)
                {
                    splitString[j] = splitString[j].TrimStart();
                }
                furnCount = splitString.Length;
                int counter = 0;
                int i = 0;
                while (counter == 0)
                {
                    string s = splitString[i];
                    GetFurnitureType getType = new GetFurnitureType(doc, _ftCounter, _ftData, s, _roomPt);

                    // Insert a family
                    FamilySymbol curFt = Utils.GetFamilySymbolByName(doc, getType.newFtFamName, getType.newFtFamType);

                    if (curFt != null)
                    {
                        curFt.Activate();

                        // Insert a family
                        FamilyInstance curFi = doc.Create.NewFamilyInstance(_roomPt, curFt, StructuralType.NonStructural);
                    }
                    i++;

                    if (i == furnCount)
                        counter = 1;
                }

            }
        }

        public class GetFurnitureType
        {
            public string curItem { get; set; }
            public int outerIndex = 1;
            public int innerIndex = 0;
            public int innerCounter = 1;
            public string curFi = "";
            public string newFtFamName = "";
            public string newFtFamType = "";
            public List<string> innerList = new List<string>();
            public GetFurnitureType(Document doc, int _ftCounter, List<List<string>> _ftData, string _curItem, XYZ _roomPt)
            {
                curItem = _curItem;
                for (int outerIndex = 1; outerIndex < _ftCounter; outerIndex++)
                {
                    innerList = _ftData[outerIndex];
                    curFi = innerList[0]; 
                    if (curFi == _curItem)

                    // Set furniture data
                    {
                        newFtFamName = innerList[1];
                        newFtFamType = innerList[2];
                        curFi = "";
                        break;
                    }
                }
            }
        }

        public class ReplaceFurniture
        {
            string curItem = "";
            public int outerIndex = 1;
            public int innerIndex = 0;
            public int innerCounter = 1;
            public string curFt = "";
            public string newFtFamName = "";
            public string newFtFamType = "";
            public List<string> innerList = new List<string>();
            public ReplaceFurniture(Document doc, int _ftCounter, List<List<string>> _ftData, string _curItem, XYZ _roomPt)
            {
                curItem = _curItem;
                for (int outerIndex = 1; outerIndex < _ftCounter; outerIndex++)
                {
                    innerList = _ftData[outerIndex];
                    curFt = innerList[0];
                    if (curFt == _curItem)

                    // Place the furniture
                    {
                        newFtFamName = innerList[1];
                        newFtFamType = innerList[2];

                        // Insert a family
                        FamilySymbol curFt = Utils.GetFamilySymbolByName(doc, newFtFamName, newFtFamType);

                        // Activate family symbol
                        curFt.Activate();

                        // Insert a family
                        FamilyInstance curFi = doc.Create.NewFamilyInstance(_roomPt, curFt, StructuralType.NonStructural);
                    }
                }
            }
        }

    }
}
