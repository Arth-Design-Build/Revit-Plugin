using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class PlaceViewport : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                var titleBlockIds = new List<ElementId>();
                var form = new System.Windows.Forms.Form();
                form.Size = new System.Drawing.Size(400, 400);
                form.AutoScroll = true;
                // Get the available title blocks in the project
                FilteredElementCollector collect = new FilteredElementCollector(doc);
                collect.OfCategory(BuiltInCategory.OST_TitleBlocks);
                var titleBlocks = collect.ToElements();

                // Create a checkbox for each title block
                int top = 20;
                foreach (var titleBlock in titleBlocks)
                {
                    //var checkBox = new System.Windows.Forms.CheckBox();
                    var radio = new System.Windows.Forms.RadioButton();
                    radio.Text = titleBlock.Name;
                    radio.AutoSize = true;
                    radio.Left = 20;
                    radio.Top = top;
                    radio.CheckedChanged += (sender, args) => {
                        var check = (System.Windows.Forms.RadioButton)sender;
                        if (check.Checked)
                        {
                            titleBlockIds.Add(titleBlock.Id);
                        }
                        else
                        {
                            titleBlockIds.Remove(titleBlock.Id);
                        }
                    };

                    form.Controls.Add(radio);
                    top += 25;
                }
                var button = new Button
                {
                    Text = "Select",
                    Left = 150,
                    Top = titleBlocks.Count * 30 + 40
                };
                button.AutoSize = true;
                form.Controls.Add(button);
                button.Click += (sender, args) =>
                {
                    form.Close();
                };
                form.ShowDialog();
                
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    // Load the Excel file
                    FileInfo file = new FileInfo(filePath);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        // Get the first worksheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                        using (Transaction transaction = new Transaction(doc, "Place Viewports"))
                        {
                            transaction.Start();

                            // Read the data from the worksheet
                            int row = 1;
                            while (worksheet.Cells[row, 1].Value != null)
                            {
                                string sheetNumber = worksheet.Cells[row, 1].Value.ToString();
                                string sheetName = worksheet.Cells[row, 2].Value.ToString();
                                string temp = worksheet.Cells[row, 3].Value.ToString();
                                char[] spearator = { ',' };
                                String[] strlist = temp.Split(spearator,StringSplitOptions.RemoveEmptyEntries);
                                //string viewName = worksheet.Cells[row, 3].Value.ToString();

                                // Create a new sheet
                                ViewSheet sheet = ViewSheet.Create(doc, titleBlockIds[0]);
                                sheet.SheetNumber = sheetNumber;
                                sheet.Name = sheetName;

                                int count = strlist.Length;
                                double num = strlist.Length+1;
                                double i = 1.00;
                                double j = num - 1;
                                double test = 1;

                                foreach (String s in strlist)
                                {
                                    double p = i / num;
                                    //TaskDialog.Show("P", p.ToString());
                                    double q = j / num;
                                    string viewName = s;
                                    //TaskDialog.Show("View", viewName);
                                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                                    collector.OfCategory(BuiltInCategory.OST_Views);
                                    collector.OfClass(typeof(Autodesk.Revit.DB.View));
                                    Autodesk.Revit.DB.View view = collector.FirstOrDefault<Element>(e => e.Name.Equals(viewName)) as Autodesk.Revit.DB.View;

                                    if (view != null)
                                    {

                                        // Place the view on the sheet
                                        //BoundingBoxUV outline = sheet.Outline - sheet.Title;
                                        BoundingBoxUV outline = sheet.Outline;
                                        double xu = (outline.Max.U + outline.Min.U);
                                        double yu = (outline.Max.V + outline.Min.V);
                                        //TaskDialog.Show(xu.ToString(), yu.ToString());
                                        double xu_new = ((p * xu) + (q * 0)) / (p + q);
                                        double yu_new = ((p * yu) + (q * 0)) / (p + q);
                                        if ((outline.Max.U + outline.Min.U)>=2.29)
                                        {
                                            view.Scale = 100 * count;
                                        }
                                        else
                                        {
                                            if((outline.Max.U + outline.Min.U) >= 1.96 && (outline.Max.U + outline.Min.U) < 2.29)
                                            {
                                                view.Scale = 125 * count;
                                            }
                                            else
                                            {
                                                if ((outline.Max.U + outline.Min.U) >= 1.64 && (outline.Max.U + outline.Min.U) < 1.96)
                                                {
                                                    view.Scale = 150 * count;
                                                }
                                                else
                                                {
                                                    if ((outline.Max.U + outline.Min.U) >= 1.31 && (outline.Max.U + outline.Min.U) < 1.64)
                                                    {
                                                        view.Scale = 175 * count;
                                                    }
                                                    else
                                                    {
                                                        if ((outline.Max.U + outline.Min.U) >= 0.98 && (outline.Max.U + outline.Min.U) < 1.31)
                                                        {
                                                            view.Scale = 200 * count;
                                                        }
                                                        else
                                                        {
                                                            if ((outline.Max.U + outline.Min.U) >= 0.65 && (outline.Max.U + outline.Min.U) < 0.98)
                                                            {
                                                                view.Scale = 225 * count;
                                                            }
                                                            else
                                                            {
                                                                if ((outline.Max.U + outline.Min.U) >= 0.32 && (outline.Max.U + outline.Min.U) < 0.65)
                                                                {
                                                                    view.Scale = 250 * count;
                                                                }
                                                                else
                                                                {
                                                                    if ((outline.Max.U + outline.Min.U) >= 0 && (outline.Max.U + outline.Min.U) < 0.32)
                                                                    {
                                                                        view.Scale = 275 * count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        XYZ point = new XYZ(xu_new, yu_new, 0);
                                        Viewport viewport = Viewport.Create(doc, sheet.Id, view.Id, point);
                                    }
                                    i++;
                                    j--;
                                }

                                row++;
                            }

                            transaction.Commit();
                        }
                    }

                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
                
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
