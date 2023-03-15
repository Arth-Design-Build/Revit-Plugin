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
using System.Drawing;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Schedule : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string xlSheetName;
            int flag = 0;
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
                form.Size = new System.Drawing.Size(700, 400);
                form.Padding = new System.Windows.Forms.Padding(10, 10, 0, 10);

                form.Text = "Export Schedules in Excel";
                form.AutoScroll = true;
                var checkBoxes = new List<CheckBox>();
                foreach (ViewSchedule vs in col)
                {
                    if(vs.Name.ToString().Split(' ')[0]=="<Revision")
                    {
                        continue;
                    }
                    schedules.Add(vs);
                    var checkBox = new CheckBox
                    {
                        Text = vs.Name,
                        AutoSize = true,
                        Left = 10,
                        Top = 10 + checkBoxes.Count * 30
                    };
                    checkBoxes.Add(checkBox);
                    form.Controls.Add(checkBox);
                }
                var button1 = new Button
                {
                    Text = "Export in Multiple Files",
                    Left = 180,
                    Top = checkBoxes.Count * 40 + 10
                };
                button1.AutoSize = true;
                button1.Click += (sender, args) =>
                {
                    var selectedSchedules = checkBoxes
                        .Where(x => x.Checked)
                        .Select(x => schedules[checkBoxes.IndexOf(x)]);
                    /*
                    foreach (var items in checkBoxes)
                    {
                        if (items.Checked)
                        {
                            flag = 1;
                        }
                    }
                    */
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
                        //excelEngine.SaveAs(new FileInfo(MapPath + docName + ".xlsx"));
                        MessageBox.Show("Exported Successfully!", "Success");
                    }
                };
                var button = new Button
                {
                    Text = "Export in Single File",
                    Left = 340,
                    Top = checkBoxes.Count * 40 + 10
                };
                button.AutoSize = true;
                button.Click += (sender, args) =>
                {
                    var selectedSchedules = checkBoxes
                        .Where(x => x.Checked)
                        .Select(x => schedules[checkBoxes.IndexOf(x)]);
                    /*
                    foreach(var items in checkBoxes)
                    {
                        if (items.Checked)
                        {
                            flag = 1;
                        }
                    }
                    */
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
                                //File.WriteAllLines(MapPath + fileName + ".csv", data, Encoding.UTF8);
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
                        /*
                        if(flag==0)
                        {
                            MessageBox.Show("Not Exported!", "Failed");
                        }
                        else
                        {
                            MessageBox.Show("Exported Successfully!", "Success");
                        }
                        */
                    }
                };
                form.Controls.Add(button1);
                form.Controls.Add(button);
                System.Windows.Forms.Panel blankPanel = new System.Windows.Forms.Panel();
                blankPanel.Height = 10;
                form.Controls.Add(blankPanel);

                // Define the border style of the form to a dialog box.
                form.FormBorderStyle = FormBorderStyle.FixedDialog;

                // Set the MaximizeBox to false to remove the maximize box.
                form.MaximizeBox = false;

                // Set the MinimizeBox to false to remove the minimize box.
                form.MinimizeBox = false;

                form.Icon = new Icon("favicon.ico");

                form.BackColor = System.Drawing.Color.LightGray;

                form.ShowDialog();
            }
            return Result.Succeeded;
        }
    }
}



