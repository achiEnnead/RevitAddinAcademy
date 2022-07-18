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
using Forms = System.Windows.Forms;

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

            //Get file with dialog box

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.Filter = "Excel Files | *.xlsx; *.xls";

            string filePath = "";

            if(dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                filePath = dialog.FileName;
            }

       
            try
            {
                //open excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWb = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet excelWsLevels = excelWb.Worksheets.Item[1];
                Excel.Worksheet excelWsSheets = excelWb.Worksheets.Item[2];

                Excel.Range excelRngLevels = excelWsLevels.UsedRange;
                Excel.Range excelRngSheets = excelWsSheets.UsedRange;

                int rowCountLevels = excelRngLevels.Rows.Count;
                int rowCountSheets = excelRngSheets.Rows.Count;

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Levels and Sheets");

                    //get view type IDs
                    FilteredElementCollector collectorViews = new FilteredElementCollector(doc);
                    collectorViews.OfClass(typeof(ViewFamilyType));

                    ViewFamilyType curVFT = null;
                    ViewFamilyType curRCPVFT = null;
                    foreach (ViewFamilyType curElem in collectorViews)
                    {
                        if (curElem.ViewFamily == ViewFamily.FloorPlan)
                        {
                            curVFT = curElem;
                        }
                        else if (curElem.ViewFamily == ViewFamily.CeilingPlan)
                        {
                            curRCPVFT = curElem;
                        }
                    }

                    //make levels & Views

                    for (int i = 2; i < rowCountLevels; i++)
                    {
                        
                        Excel.Range levelData1 = excelWsLevels.Cells[i, 1];
                        Excel.Range levelData2 = excelWsLevels.Cells[i, 2];

                        LeveLStruct lStruct = new LeveLStruct(levelData1.Value.ToString(), levelData2.Value);

                        //string levelName = levelData1.Value.ToString();
                        //double levelElev = levelData2.Value;

                        Level newLevel = Level.Create(doc, levelElev);
                        newLevel.Name = levelName;

                        ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                        ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, newLevel.Id);
                        curRCP.Name = curRCP.Name + " RCP";
                    }

                    //make sheets

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    collector.WhereElementIsElementType();

                    for (int j = 2; j < rowCountSheets; j++)
                    {
                        Excel.Range sheetData1 = excelWsSheets.Cells[j, 1];
                        Excel.Range sheetData2 = excelWsSheets.Cells[j, 2];
                        Excel.Range sheetData3 = excelWsSheets.Cells[j, 3];
                        Excel.Range sheetData4 = excelWsSheets.Cells[j, 4];
                        Excel.Range sheetData5 = excelWsSheets.Cells[j, 5];

                        string sheetNum = sheetData1.Value.ToString();
                        string sheetName = sheetData2.Value.ToString();
                        string sheetView = sheetData3.Value.ToString();
                        string sheetDrawnBy = sheetData4.Value.ToString();
                        string sheetCheckedBy = sheetData5.Value.ToString();

                        ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());
                        newSheet.SheetNumber = sheetNum;
                        newSheet.Name = sheetName;

                        //put View on Sheet if names match

                        View existingView = GetViewByName(doc, sheetView);

                        if(existingView != null)
                        {
                            Viewport newVP = Viewport.Create(doc, newSheet.Id, existingView.Id, new XYZ(0, 0, 0));
                        }
                        else
                        {
                            TaskDialog.Show("Error", "Could not find view" + sheetView);
                        }

                        //Drawn By and Checked by
                        foreach (Parameter curParam in newSheet.Parameters)
                        {
                            if (curParam.Definition.Name == "Drawn By")
                            {
                                curParam.Set(sheetDrawnBy);
                            }
                        }

                        foreach (Parameter curParam in newSheet.Parameters)
                        {
                            if (curParam.Definition.Name == "Checked By")
                            {
                                curParam.Set(sheetCheckedBy);
                            }
                        }
                    }



                    

                    t.Commit();
                }

                excelWb.Close();
                excelApp.Quit();
            }
            catch(Exception ex)
            {

                Debug.Print(ex.Message);
            }



            return Result.Succeeded;
        }

        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collectorView = new FilteredElementCollector(doc);
            collectorView.OfClass(typeof(View));

            foreach(View curView in collectorView)
            {
                if(curView.Name == viewName)
                {
                    return curView;
                }
            }
            return null;          
                
        }

        internal struct LeveLStruct
        {
            public string Name;
            public double Elevation;
    
            public LeveLStruct(string name, double elevation)
            {
                Name = name;
                Elevation = elevation;
            }
        }

        internal struct SheetStruct
        {
            public string SheetNum;
            public string SheetName;
            public string SheetView;
            public string SheetDrawnBy;
            public string SheetCheckedBy;

            public SheetStruct(string sheetNum, string sheetName, string sheetView, string sheetDrawnBy, string sheetCheckedBy)
            {
                SheetNum = sheetNum;
                SheetName = sheetName;
                SheetView = sheetView;
                SheetDrawnBy = sheetDrawnBy;
                SheetCheckedBy = sheetCheckedBy;
            }
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
