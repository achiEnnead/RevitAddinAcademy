#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Marquee Select Elements");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            WallType curStorefrontType = GetWallTypeByName(doc, "Storefront");
            Level curLevel = GetLevelByName(doc, "Level 1");
            DuctType curDuctType = GetDuctTypeByName(doc, "Default");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");
            MEPSystemType curDomHotWater = GetMEPSysTypeByName(doc, "Domestic Hot Water");
            MEPSystemType curSupplyAir = GetMEPSysTypeByName(doc, "Supply Air");

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create elements from model lines by type");

                foreach (Element elem in pickList)
                {
                    if (elem is ModelLine)
                    {
                        ModelLine mCurve = (ModelLine)elem;
                        curveList.Add(mCurve);

                        GraphicsStyle curGS = mCurve.LineStyle as GraphicsStyle;

                        Curve curCurve = mCurve.GeometryCurve;
                        XYZ startPoint = curCurve.GetEndPoint(0);
                        XYZ endPoint = curCurve.GetEndPoint(1);

                        //try
                        //{
                        //    XYZ startPoint = curCurve.GetEndPoint(0);
                        //    XYZ endPoint = curCurve.GetEndPoint(1);
                        //    if(startPoint != null)
                        //    {
                        //        return startPoint;
                        //        return endPoint;
                        //    }

                        //}
                        //catch
                        //{
                        //    TaskDialog.Show("Error", $"Curve {curCurve.Id} has no start or end point");
                        //}



                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Wall newStorefrontWall = Wall.Create(doc, curCurve, curStorefrontType.Id, curLevel.Id, 15, 0, false, false);
                                break;
                            case "A-WALL":
                                Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);
                                break;
                            case "M-DUCT":
                                Duct newDuct = Duct.Create(
                                    doc,
                                    curSupplyAir.Id,
                                    curDuctType.Id,
                                    curLevel.Id,
                                    startPoint,
                                    endPoint);
                                break;
                            case "P-PIPE":
                                Pipe newPipe = Pipe.Create(
                                    doc,
                                    curDomHotWater.Id,
                                    curPipeType.Id,
                                    curLevel.Id,
                                    startPoint,
                                    endPoint);
                                break;
                            default:
                                break;
                        }
                    }
                }

                t.Commit();

            }



            TaskDialog.Show("Complete", curveList.Count.ToString());

            return Result.Succeeded;
        }

        private WallType GetWallTypeByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element curElem in collector)
            {
                WallType curType = curElem as WallType;

                if (curType.Name == v)
                    return curType;
            }
            return null;
        }

        private Level GetLevelByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level curType = curElem as Level;

                if (curType.Name == v)
                    return curType;
            }
            return null;
        }

        private DuctType GetDuctTypeByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element curElem in collector)
            {
                DuctType curType = curElem as DuctType;

                if (curType.Name == v)
                    return curType;
            }
            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curElem in collector)
            {
                PipeType curType = curElem as PipeType;

                if (curType.Name == v)
                    return curType;
            }
            return null;
        }

        private MEPSystemType GetMEPSysTypeByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in collector)
            {
                MEPSystemType curType = curElem as MEPSystemType;

                if (curType.Name == v)
                    return curType;
            }
            return null;
        }
    }
}
