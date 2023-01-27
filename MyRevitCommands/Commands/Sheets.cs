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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Sheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (Transaction trans = new Transaction(doc, "Create Sheet"))
                {
                    trans.Start();
                    string[] sName = new string[100];
                    string[] sTags = new string[100];
                    string[] sTitleblock = new string[100];
                    // string[] sNumbers = { "J101", "J102", "J103", "J104" };
                    int count1 = 0;
                    int count2 = 0;
                    int count3 = 0;
                    TaskDialog myDialog = new TaskDialog("Create Sheet");
                    myDialog.MainInstruction = "Select the Appropriate Titleblock";
                    // myDialog.MainContent = "Please select one of the following options:";

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfClass(typeof(FamilySymbol));

                    List<string> titleblockNames = new List<string>();
                    foreach (FamilySymbol symbol in collector)
                    {
                        if (symbol.Name.ToLower().Contains("horizontal") || symbol.Name.ToLower().Contains("vertical") || symbol.Name.ToLower().Contains("metric"))
                        {
                            titleblockNames.Add(symbol.Name.ToString());
                        }
                    }
                    for(int i=0;i<titleblockNames.Count;i++)
                    {
                        myDialog.AddCommandLink((TaskDialogCommandLinkId)(i + 1), titleblockNames[i]);
                    }

                    TaskDialogResult result = myDialog.Show();
                    int selectedOption = (int)result - 1;

                    if (selectedOption >= 0 && selectedOption < titleblockNames.Count)
                    {
                        OpenFileDialog openFileDialog1 = new OpenFileDialog();
                        // Set filter options and filter index.
                        openFileDialog1.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
                        openFileDialog1.FilterIndex = 1;
                        openFileDialog1.Multiselect = false;
                        // Call the ShowDialog method to show the dialog box.
                        var userClickedOK = openFileDialog1.ShowDialog();
                        // Process input if the user clicked OK.
                        if (userClickedOK == DialogResult.OK && openFileDialog1.FileName != null)
                        {
                            string filePath = openFileDialog1.FileName;
                            if (File.Exists(filePath))
                            {
                                FileInfo file = new FileInfo(filePath);
                                using (ExcelPackage package = new ExcelPackage(file))
                                {
                                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                                    int rowCount = worksheet.Dimension.Rows;
                                    int ColCount = worksheet.Dimension.Columns;
                                    for (int row = 1; row <= rowCount; row++)
                                    {
                                        for (int col = 1; col <= 1; col++)
                                        {
                                            // Console.Write(worksheet.Cells[row, col].Value.ToString() + " ");
                                            string temp1 = worksheet.Cells[row, col].Value.ToString();
                                            sTags[count1] = temp1;
                                            count1++;
                                        }
                                        // Console.WriteLine();
                                    }
                                    for (int row = 1; row <= rowCount; row++)
                                    {
                                        for (int col = 2; col <= 2; col++)
                                        {
                                            // Console.Write(worksheet.Cells[row, col].Value.ToString() + " ");
                                            string temp2 = worksheet.Cells[row, col].Value.ToString();
                                            sName[count2] = temp2;
                                            count2++;
                                        }
                                        // Console.WriteLine();
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("File not found at: " + filePath);
                            }
                        }
                    }
                    int l1 = sName.Length;
                    FamilySymbol tBlock = new FilteredElementCollector(doc)
                           .OfCategory(BuiltInCategory.OST_TitleBlocks)
                           .WhereElementIsElementType()
                           .Cast<FamilySymbol>()
                           .First(q => q.Name == titleblockNames[selectedOption]);
                    for (int i=0;i<count1; i++)
                    {

                        // ViewSheet vSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
                        ViewSheet vSheet = ViewSheet.Create(doc, tBlock.Id);
                        vSheet.Name = sName[i];
                        vSheet.SheetNumber = sTags[i];
                        /*
                        ViewSheet titleBlockView = new FilteredElementCollector(doc)
                                                        .OfClass(typeof(ViewSheet))
                                                        .Cast<ViewSheet>()
                                                        .FirstOrDefault(q => q.Name == "E1 30x42 Horizontal");

                        if (titleBlockView != null)
                        {
                            Viewport vp = Viewport.Create(doc, vSheet.Id, titleBlockView.Id, new XYZ(0, 0, 0));
                            doc.PrintManager.PrintRange = PrintRange.Select;
                            doc.PrintManager.PrintToFile = true;
                            doc.PrintManager.PrintToFileName = doc.PathName;
                            doc.PrintManager.Apply();
                            doc.Regenerate();

                        }
                        */
                        /*
                        FamilySymbol customTitleBlock = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_TitleBlocks)
                            .WhereElementIsElementType()
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(q => q.Name == "Custom Title Block");

                        if (customTitleBlock != null)
                        {
                            Viewport vp = Viewport.Create(doc, vSheet.Id, customTitleBlock.Id, new XYZ(0, 0, 0));
                        }
                        */

                    }
                    trans.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message= e.Message;
                return Result.Failed;
            }
            
        }
    }
}
