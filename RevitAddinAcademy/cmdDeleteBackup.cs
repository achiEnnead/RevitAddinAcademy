#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDeleteBackup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // set variables
            int counter = 0;
            string logPath = "";

            //create list for log file & list of file sizes
            List<string> deletedFileLog = new List<string>();
            deletedFileLog.Add("The following backup files have been deleted:");
            deletedFileLog.Add("Placeholder");

            //List<long> deletedFilesSizes = new List<long>();
            long deletedFileSumSize = 0;

            FolderBrowserDialog selectFolder = new FolderBrowserDialog();
            selectFolder.ShowNewFolderButton = false;

            //open folder dialog and only run code if a folder is selected
            if (selectFolder.ShowDialog() == DialogResult.OK)
            {
                //get selected folder path
                string directory = selectFolder.SelectedPath;

                //get all files from selected folder
                string[] files = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);

                //loop through files
                foreach (string file in files)
                {
                    //check if file is a revit file
                    if (Path.GetExtension(file) == ".rvt" || Path.GetExtension(file) == ".rfa")
                    {
                        string checkString = file.Substring(file.Length - 9, 9);

                        if (checkString.Contains(".00") == true)
                        {
                            //create variable FileInfo to get file size, get file size and convert to string
                            FileInfo fi = new FileInfo(file);
                            long size = fi.Length;
                            deletedFileSumSize = deletedFileSumSize + size;

                            //add filename to list
                            deletedFileLog.Add(file);

                            //delete file
                            File.Delete(file);

                            //increment counter
                            counter++;
                        }
                    }
                }

                //check digits in deletedFileSumSize
                string reportSumUnit = "";
                long finalFileSumSize = 0;

                if (deletedFileSumSize > 1024000000)
                {
                    reportSumUnit = " GB";
                    finalFileSumSize = deletedFileSumSize / 1024000000;
                }
                else if (deletedFileSumSize > 1024000)
                {
                    reportSumUnit = " MB";
                    finalFileSumSize = deletedFileSumSize / 1024000;
                }
                else if (deletedFileSumSize > 1024)
                {
                    reportSumUnit = " KB";
                    finalFileSumSize = deletedFileSumSize / 1024;
                }
                else
                {
                    reportSumUnit = " Bytes";
                    finalFileSumSize = deletedFileSumSize;
                }



                // output log file
                if (counter > 0)
                {
                    deletedFileLog[1] = ("Total file size saved: " + finalFileSumSize.ToString() + reportSumUnit);
                    logPath = WriteListToTxt(deletedFileLog, directory);
                }

                // alert user
                TaskDialog td = new TaskDialog("Complete");
                td.MainInstruction = "Deleted" + counter.ToString() + "backup files.";
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Click to view log file");
                td.CommonButtons = TaskDialogCommonButtons.Ok;

                TaskDialogResult result = td.Show();

                if (result == TaskDialogResult.CommandLink1)
                {
                    Process.Start(logPath);
                }
            }
            else
            {

            }


           
            return Result.Succeeded;
        }

        internal string WriteListToTxt(List<string> stringList, string filePath)
        {
            string fileName = "_Delete Backup Files.txt";
            string fullPath = filePath + @"\" + fileName;

            File.WriteAllLines(fullPath, stringList);

            return fullPath;
        }



    }
}
