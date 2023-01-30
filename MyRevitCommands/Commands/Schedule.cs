using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Globalization;
using System.Reflection;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Schedule : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string xlSheetName;
            var folderBrowser = new FolderBrowserDialog();
            folderBrowser.Description = "Select Folder to Save the Exported Schedules";
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                var MapPath = folderBrowser.SelectedPath + "\\";
                var uidoc = commandData.Application.ActiveUIDocument;
                var doc = uidoc.Document;
                var docName = doc.Title;
                var col = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule));
                var opt = new ViewScheduleExportOptions
                {
                    TextQualifier = ExportTextQualifier.None,
                    FieldDelimiter = ",",
                };
                var schedules = new List<ViewSchedule>();
                var form = new System.Windows.Forms.Form();
                form.Size = new System.Drawing.Size(400, 400);
                var checkBoxes = new List<CheckBox>();
                foreach (ViewSchedule vs in col)
                {
                    schedules.Add(vs);
                    var checkBox = new CheckBox
                    {
                        Text = vs.Name,
                        AutoSize = true,
                        Left = 10,
                        Top = 10 + checkBoxes.Count * 40
                    };
                    checkBoxes.Add(checkBox);
                    form.Controls.Add(checkBox);
                }
                var button = new Button
                {
                    Text = "Export",
                    Left = 150,
                    Top = checkBoxes.Count * 40 + 10
                };
                button.Click += (sender, args) =>
                {
                var selectedSchedules = checkBoxes
                    .Where(x => x.Checked)
                    .Select(x => schedules[checkBoxes.IndexOf(x)]);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excelEngine = new ExcelPackage())
                {
                    foreach (ViewSchedule vs in selectedSchedules)
                    {
                        var fileName = vs.Name;
                        if (vs.Name != null)
                        {
                            if (vs.Name.Length > 30)
                                xlSheetName = vs.Name.Substring(0, 30);
                            else
                                xlSheetName = vs.Name;
                            var ws = excelEngine.Workbook.Worksheets.Add(xlSheetName);
                            vs.Export(MapPath, fileName + ".txt", opt);
                            var data = File.ReadAllLines(MapPath + fileName + ".txt");
                            File.WriteAllLines(MapPath + fileName + ".csv", data, Encoding.UTF8);
                            File.Delete(MapPath + fileName + ".txt");
                            for (int i = 0; i < data.Count(); i++)
                            {
                                string[] fields = data.ElementAt(i).Split(',');
                                for (int j = 0; j < fields.Length; j++)
                                {
                                    string field = fields[j];
                                    if (field.StartsWith("\"") && field.EndsWith("\""))
                                    {
                                        field = field.Substring(1, field.Length - 2);
                                    }
                                    ws.Cells[i + 1, j + 1].Value = field;
                                }
                            }
                        }
                    }
                    excelEngine.SaveAs(new FileInfo(MapPath + docName+ ".xlsx"));
                    MessageBox.Show("Exported Successfully!", "Success");
                }
                };
                form.Controls.Add(button);
                form.ShowDialog();
            }
            return Result.Succeeded;
        }
    }
}



