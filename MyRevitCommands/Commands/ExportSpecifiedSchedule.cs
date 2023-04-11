using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class ExportSpecifiedSchedule : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string xlSheetName;
            var folderBrowser = new FolderBrowserDialog();
            folderBrowser.Description = "Select Folder to Save the Exported Schedules";
            if (true)
            {
                var exportPath = "C:\\Users\\ASUS\\Downloads" + "\\";
                var uidoc = commandData.Application.ActiveUIDocument;
                var doc = uidoc.Document;
                var docName = doc.Title;

                // Read the schedule names from the specified Excel file
                var scheduleNames = new List<string>();
                var excelFile = new FileInfo(Path.Combine(exportPath, "ScheduleNames.xlsx"));
                if (excelFile.Exists)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var excelPackage = new ExcelPackage(excelFile))
                    {
                        var worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet != null)
                        {
                            var row = 1;
                            var nameCell = worksheet.Cells[row, 1];
                            while (!string.IsNullOrEmpty(nameCell.Text))
                            {
                                scheduleNames.Add(nameCell.Text);
                                row++;
                                nameCell = worksheet.Cells[row, 1];
                            }
                        }
                    }
                }

                if (scheduleNames.Count == 0)
                {
                    TaskDialog.Show("Export Schedules", "No Schedule Names Found in the Specified Excel File.");
                    return Result.Cancelled;
                }

                var schedules = new List<ViewSchedule>();
                foreach (var name in scheduleNames)
                {
                    var schedule = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(vs => vs.Name == name);
                    if (schedule != null)
                    {
                        schedules.Add(schedule);
                    }
                    else
                    {
                        TaskDialog.Show("Export Schedules", $"Schedule '{name}' Not Found in the Document.");
                    }
                }

                if (schedules.Count == 0)
                {
                    TaskDialog.Show("Export Schedules", "No Schedules Found in the Document Matching the Specified Names");
                    return Result.Cancelled;
                }

                var opt = new ViewScheduleExportOptions
                {
                    TextQualifier = ExportTextQualifier.None,
                    FieldDelimiter = ",",
                };

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excelPackage = new ExcelPackage())
                {
                    foreach (var schedule in schedules)
                    {
                        var fileName = schedule.Name;
                        if (fileName.Length > 30)
                            xlSheetName = fileName.Substring(0, 30);
                        else
                            xlSheetName = fileName;

                        var worksheet = excelPackage.Workbook.Worksheets.Add(xlSheetName);
                        schedule.Export(exportPath, fileName + ".txt", opt);
                        var data = File.ReadAllLines(Path.Combine(exportPath, fileName + ".txt"));
                        File.WriteAllLines(Path.Combine(exportPath, fileName + ".csv"), data, Encoding.UTF8);
                        File.Delete(Path.Combine(exportPath, fileName + ".txt"));
                        for (int i = 0; i < data.Count(); i++)
                        {
                            worksheet.Cells[i + 1, 1].Value = data[i];
                        }
                        // Autofit columns
                        worksheet.Cells.AutoFitColumns();
                    }

                    var exportFile = new FileInfo(Path.Combine(exportPath, $"{docName}_Schedules.xlsx"));
                    excelPackage.SaveAs(exportFile);
                }

                TaskDialog.Show("Export Schedules", "Schedules Exported Successfully!");
                return Result.Succeeded;
            }

            return Result.Cancelled;
        }
    }
}