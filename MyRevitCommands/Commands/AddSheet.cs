using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddSheet : IExternalCommand
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
                var form = new System.Windows.Forms.Form
                {
                    Size = new System.Drawing.Size(400, 400),
                    AutoScroll = true
                };
                // Get the available title blocks in the project
                FilteredElementCollector collect = new FilteredElementCollector(doc);
                collect.OfCategory(BuiltInCategory.OST_TitleBlocks);
                var titleBlocks = collect.ToElements();

                IEnumerable<Element> uniqueTitleBlocks = titleBlocks.Cast<Element>().GroupBy(x => x.Name).Select(x => x.First());

                // Create a checkbox for each title block
                int top = 20;
                var temp = "";
                int itr = 0;
                int f1 = 0;
                int f2 = 0;
                foreach (var titleBlock in uniqueTitleBlocks)
                {
                    if (titleBlock.Name == temp)
                    {
                        continue;
                    }
                    //var checkBox = new System.Windows.Forms.CheckBox();
                    var radio = new System.Windows.Forms.RadioButton
                    {
                        Text = titleBlock.Name
                    };
                    temp = titleBlock.Name;
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
                    itr++;
                }
                var button = new Button
                {
                    Text = "Select",
                    Left = 150,
                    Top = itr * 30 + 20,
                    AutoSize = true
                };
                form.Controls.Add(button);
                button.Click += (sender, args) =>
                {
                    f1 = 1;
                    form.Close();
                };
                form.Text = "Add Sheet Data";

                form.AutoScroll = true;

                // Define the border style of the form to a dialog box.
                form.FormBorderStyle = FormBorderStyle.FixedDialog;

                // Set the MaximizeBox to false to remove the maximize box.
                form.MaximizeBox = false;

                // Set the MinimizeBox to false to remove the minimize box.
                form.MinimizeBox = false;

                string iconUrl = "https://www.linkpicture.com/q/favicon_77.ico";
                WebRequest request = WebRequest.Create(iconUrl);
                using (WebResponse response = request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    if (memoryStream != null)
                    {
                        form.Icon = new Icon(memoryStream);
                    }
                    else
                    {
                        // Handle the case where the memory stream is null
                    }
                }

                form.BackColor = System.Drawing.Color.LightGray;

                form.ShowDialog();

                if (f1 == 1)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                    };
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

                            using (Transaction transaction = new Transaction(doc, "Create Sheets"))
                            {
                                transaction.Start();

                                int firstrow = 1;
                                int colcount = 1;
                                int sheetnocol = 0;
                                int sheetnamecol = 0;
                                int clientnamecol = 0;
                                int projectnamecol = 0;
                                int projectnocol = 0;
                                int issuedatecol = 0;
                                int drawnbycol = 0;
                                int checkedbycol = 0;

                                while (worksheet.Cells[firstrow, colcount].Value != null)
                                {
                                    string value = worksheet.Cells[firstrow, colcount].Value.ToString();
                                    if (value == "Sheet Number" || value == "Sheet No" || value == "Sheet No." || value == "sheet number" || value == "sheet no" || value == "sheet no." || value == "Sheet number" || value == "Sheet no" || value == "Sheet no." || value == "sheet Number" || value == "sheet No" || value == "sheet No.")
                                    {
                                        sheetnocol = colcount;
                                    }
                                    else
                                    {
                                        if (value == "Sheet Name" || value == "Sheet name" || value == "sheet name" || value == "sheet Name")
                                        {
                                            sheetnamecol = colcount;
                                        }
                                        else
                                        {
                                            if (value == "Client Name" || value == "client name" || value == "Client name" || value == "client Name")
                                            {
                                                clientnamecol = colcount;
                                            }
                                            else
                                            {
                                                if (value == "Project Name" || value == "Project name" || value == "project name" || value == "project Name")
                                                {
                                                    projectnamecol = colcount;
                                                }
                                                else
                                                {
                                                    if (value == "Project Number" || value == "Project No" || value == "Project No." || value == "Project number" || value == "Project no" || value == "Project no." || value == "project number" || value == "project no" || value == "project no." || value == "project Number" || value == "project No" || value == "project No.")
                                                    {
                                                        projectnocol = colcount;
                                                    }
                                                    else
                                                    {
                                                        if (value == "Issue Date" || value == "issue date" || value == "Issue date" || value == "issue Date")
                                                        {
                                                            issuedatecol = colcount;
                                                        }
                                                        else
                                                        {
                                                            if (value == "Drawn by" || value == "Drawn By" || value == "drawn by" || value == "drawn By")
                                                            {
                                                                drawnbycol = colcount;
                                                            }
                                                            else
                                                            {
                                                                if (value == "Checked By" || value == "Checked by" || value == "checked by" || value == "checked By")
                                                                {
                                                                    checkedbycol = colcount;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    colcount++;
                                }

                                int row = 2;
                                while (worksheet.Cells[row, 1].Value != null)
                                {
                                    string sheetNumber = null;
                                    string sheetName = null;
                                    string clientName = null;
                                    string projectName = null;
                                    string projectNumber = null;
                                    string date = null;
                                    string drawnBy = null;
                                    string checkedBy = null;
                                    if (sheetnocol != 0)
                                    {
                                        //TaskDialog.Show("Sheet No. Col", sheetnocol.ToString());
                                        if (worksheet.Cells[row, sheetnocol].Value == null)
                                        {
                                            sheetNumber = " ";
                                        }
                                        else
                                        {
                                            sheetNumber = worksheet.Cells[row, sheetnocol].Value.ToString();
                                        }
                                    }

                                    if (sheetnamecol != 0)
                                    {
                                        //TaskDialog.Show("Sheet Name Col", sheetnamecol.ToString());
                                        if (worksheet.Cells[row, sheetnamecol].Value == null)
                                        {
                                            sheetName = "Unnamed";
                                        }
                                        else
                                        {
                                            sheetName = worksheet.Cells[row, sheetnamecol].Value.ToString();
                                        }
                                    }

                                    if (clientnamecol != 0)
                                    {
                                        //TaskDialog.Show("Client Name Col", clientnamecol.ToString());
                                        if (worksheet.Cells[row, clientnamecol].Value == null)
                                        {
                                            clientName = clientName;   
                                        }
                                        else
                                        {
                                            clientName = worksheet.Cells[row, clientnamecol].Value.ToString();
                                        }
                                    }

                                    if (projectnamecol != 0)
                                    {
                                        //TaskDialog.Show("Project Name Col", projectnamecol.ToString());
                                        if (worksheet.Cells[row, projectnamecol].Value == null)
                                        {
                                            projectName = projectName;
                                        }
                                        else
                                        {
                                            projectName = worksheet.Cells[row, projectnamecol].Value.ToString();
                                        }
                                    }

                                    if (projectnocol != 0)
                                    {
                                        //TaskDialog.Show("Project No. Col", projectnocol.ToString());
                                        if (worksheet.Cells[row, projectnocol].Value == null)
                                        {
                                            projectNumber = projectNumber;
                                        }
                                        else
                                        {
                                            projectNumber = worksheet.Cells[row, projectnocol].Value.ToString();
                                        }
                                    }

                                    if (issuedatecol != 0)
                                    {
                                        //TaskDialog.Show("Issue Date Col", issuedatecol.ToString());
                                        if (worksheet.Cells[row, issuedatecol].Value == null)
                                        {
                                            date = date;
                                        }
                                        else
                                        {
                                            date = worksheet.Cells[row, issuedatecol].Value.ToString();
                                        }
                                    }

                                    if (drawnbycol != 0)
                                    {
                                        //TaskDialog.Show("Drawn By Col", drawnbycol.ToString());
                                        if (worksheet.Cells[row, drawnbycol].Value == null)
                                        {
                                            drawnBy = "Unnamed";
                                        }
                                        else
                                        {
                                            drawnBy = worksheet.Cells[row, drawnbycol].Value.ToString();
                                        }
                                    }

                                    if (checkedbycol != 0)
                                    {
                                        //TaskDialog.Show("Checked By Col", checkedbycol.ToString());
                                        if (worksheet.Cells[row, checkedbycol].Value == null)
                                        {
                                            checkedBy = "Unnamed";
                                        }
                                        else
                                        {
                                            checkedBy = worksheet.Cells[row, checkedbycol].Value.ToString();
                                        }
                                    }

                                    // Create a new sheet
                                    ViewSheet sheet = ViewSheet.Create(doc, titleBlockIds[0]);

                                    ProjectInfo projectInfo = doc.ProjectInformation;

                                    if (clientName != null)
                                    {
                                        projectInfo.ClientName = clientName;
                                    }

                                    if (projectName != null)
                                    {
                                        projectInfo.Name = projectName;
                                    }

                                    if (projectNumber != null)
                                    {
                                        projectInfo.Number = projectNumber;
                                    }

                                    if (date != null)
                                    {
                                        projectInfo.IssueDate = date;
                                    }

                                    Parameter sheetNoParam = sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER);
                                    if (sheetNoParam != null && sheetNumber != null)
                                    {
                                        sheetNoParam.Set(sheetNumber);
                                    }

                                    Parameter sheetNameParam = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                                    if (sheetNameParam != null && sheetName != null)
                                    {
                                        sheetNameParam.Set(sheetName);
                                    }

                                    Parameter drawnByParam = sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY);
                                    if (drawnByParam != null && drawnBy != null)
                                    {
                                        drawnByParam.Set(drawnBy);
                                    }

                                    Parameter checkedByParam = sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY);
                                    if (checkedByParam != null && checkedBy != null)
                                    {
                                        checkedByParam.Set(checkedBy);
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
