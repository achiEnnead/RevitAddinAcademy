#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Session02Challenge: IExternalCommand
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

            string excelFile = @"C:\temp\Session02_Challenge.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWsLevels = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWsSheets = excelWb.Worksheets.Item[2];

            Excel.Range rngLevels = excelWsLevels.UsedRange;
            int rowCountLevels = rngLevels.Rows.Count;

            Excel.Range rngSheets = excelWsSheets.UsedRange;
            int rowCountSheets = rngSheets.Rows.Count;

            List<string[]> dataListLevels = new List<string[]>();
            List<string[]> dataListSheets = new List<string[]>();

            for (int i = 1; i < rowCountLevels; i++)
            {
                Excel.Range cell1 = excelWsLevels.Cells[i, 1];
                Excel.Range cell2 = excelWsLevels.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell1.Value.ToString();

                string[] dataArrayLevels = new string[3];
                dataArrayLevels[0] = data1;
                dataArrayLevels[1] = data2;
            }

            for (int i = 1; i < rowCountSheets; i++)
            {
                Excel.Range cell1 = excelWsLevels.Cells[i, 1];
                Excel.Range cell2 = excelWsLevels.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell1.Value.ToString();
       
                string[] dataArraySheets = new string[2];
                dataArraySheets[0] = data1;
                dataArraySheets[1] = data2;
             }

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels and Sheets");    

                foreach (int i in dataListLevels)
                {
                    Level curLevel = Level.Create(doc, i[1]);
                    curLevel.Name = i[0];
                }



                t.Commit();
            }

            excelWb.Close();
            excelApp.Quit();


            return Result.Succeeded;
        }

        ////method for getting total used rows and columns
        //internal int GetRowCount(Excel.Worksheet ws)
        //{
        //    Excel.Range excelRng = ws.UsedRange;
        //    Result = ws.Rows.Count;
        //}

        ////method for adding all data into the list
        //internal string[] DataArray(int rowCount, int columnCount, Excel.Worksheet ws);
    }
}
