#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command01Challenge : IExternalCommand
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

            string output1 = "FIZZ";
            string output2 = "BUZZ";
            int range = 100;
            List<string> resultsList = new List<string>();

            string fileName = doc.PathName;

            double offset = 0.05;
            double offsetCalc = offset * doc.ActiveView.Scale;

            XYZ currentPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);

            FilteredElementCollector collector = new FilteredElementCollector(doc);

            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note");
            t.Start();

            for (int i = 1; i <= range; i++)
            {
                if (i % 3 == 0 && i % 5 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, currentPoint, output1 + output2, collector.FirstElementId());
                }
                else if (i % 3 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, currentPoint, output1, collector.FirstElementId());
                }
                else if (i % 5 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, currentPoint, output2, collector.FirstElementId());
                }
                else
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, currentPoint, i.ToString(), collector.FirstElementId());
                }
                currentPoint = currentPoint.Subtract(offsetPoint);
            }


            t.Commit();
            t.Dispose();
        



                return Result.Succeeded;
        }
    }
}
