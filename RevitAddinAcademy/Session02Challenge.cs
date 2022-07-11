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

            //open excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWs2 = excelWb.Worksheets.Item[2];

            Excel.Range excelRng1 = excelWs1.UsedRange;
            Excel.Range excelRng2 = excelWs2.UsedRange;

            int rowCountLevels = excelRng2.Rows.Count;
            int rowCountSheets = excelRng2.Rows.Count;

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels and Sheets");

                for (int i = 2; i < rowCountLevels; i++)
                {
                    Excel.Range levelData1 = excelWs1.Cells[i, 1];
                    Excel.Range levelData2 = excelWs1.Cells[i, 2];

                    string levelName = levelData1.Value.ToString();
                    double levelElev = levelData2.Value;

                    Level newLevel = Level.Create(doc, levelElev);
                    newLevel.Name = levelName;
                }

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                for (int j = 2; j < rowCountSheets; j++)
                {
                    Excel.Range sheetData1 = excelWs1.Cells[j, 1];
                    Excel.Range sheetData2 = excelWs1.Cells[j, 2];

                    string sheetNum = sheetData1.Value.ToString();
                    string sheetName = sheetData2.Value.ToString();

                    ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    newSheet.SheetNumber = sheetNum;
                    newSheet.Name = sheetName;
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
