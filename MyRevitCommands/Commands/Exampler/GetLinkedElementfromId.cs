using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Windows.Forms;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class GetLinkedElementfromId : IExternalCommand
    {
        private System.Windows.Forms.Form _form;
        private Label _elementIdLabel;
        private Button _copyButton;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the active Revit document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Autodesk.Revit.DB.Document doc = uidoc.Document;

                // Create a custom dialog box to prompt the user for the element ID
                using (var form = new System.Windows.Forms.Form())
                {
                    var label = new System.Windows.Forms.Label();
                    label.Text = "Enter the Element ID:";
                    label.Location = new System.Drawing.Point(10, 10);
                    label.AutoSize = true;
                    form.Controls.Add(label);

                    var textBox = new System.Windows.Forms.TextBox();
                    textBox.Location = new System.Drawing.Point(10, 35);
                    textBox.Width = 250;
                    form.Controls.Add(textBox);

                    var okButton = new System.Windows.Forms.Button();
                    okButton.Text = "Ok";
                    okButton.Location = new System.Drawing.Point(60, 65);
                    okButton.Click += (sender, e) => { form.DialogResult = System.Windows.Forms.DialogResult.OK; };
                    form.Controls.Add(okButton);

                    var cancelButton = new System.Windows.Forms.Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.Location = new System.Drawing.Point(135, 65);
                    cancelButton.Click += (sender, e) => { form.DialogResult = System.Windows.Forms.DialogResult.Cancel; };
                    form.Controls.Add(cancelButton);

                    form.AcceptButton = okButton;
                    form.CancelButton = cancelButton;
                    form.ClientSize = new System.Drawing.Size(270, 100);
                    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

                    // Display the dialog box and get the result
                    System.Windows.Forms.DialogResult result = form.ShowDialog();

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        // Parse the input as an element ID
                        string input = textBox.Text;
                        if (string.IsNullOrEmpty(input))
                        {
                            return Result.Cancelled;
                        }
                        //TaskDialog.Show("Input", input.ToString());
                        ElementId linkedElementId = new ElementId(int.Parse(input));
                        //TaskDialog.Show("Id", linkedElementId.ToString());

                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        Color red = new Color(255, 0, 0);
                        Element solidFill = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Where(q => q.Name.Contains("Solid")).First();

                        ogs.SetProjectionLineColor(red);
                        ogs.SetProjectionLineWeight(4);

                        Autodesk.Revit.DB.Document linkedDoc = null;
                        FilteredElementCollector rvtLinks = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(typeof(RevitLinkType));
                        if (rvtLinks.ToElements().Count > 0)
                        {
                            foreach (RevitLinkType rvtLink in rvtLinks.ToElements())
                            {
                                if (rvtLink.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                                {
                                    RevitLinkInstance link = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(typeof(RevitLinkInstance)).Where(x => x.GetTypeId() == rvtLink.Id).First() as RevitLinkInstance;
                                    linkedDoc = link.GetLinkDocument();
                                }
                            }
                        }
                        //TaskDialog.Show("Linked Element", linkedDoc.Title.ToString());

                        Element linkedElement = linkedDoc.GetElement(linkedElementId);
                        FilteredElementCollector linkedViews = new FilteredElementCollector(linkedDoc).OfClass(typeof(Autodesk.Revit.DB.View));
                        Autodesk.Revit.DB.View linkedView = linkedViews
                        .OfType<Autodesk.Revit.DB.ViewPlan>()
                        .FirstOrDefault(v => v.Name == "Level 1");

                        //Autodesk.Revit.DB.View linkedView = linkedViews.Cast<Autodesk.Revit.DB.View>().Skip(0).FirstOrDefault();

                        //TaskDialog.Show("Linked View", linkedView.Title.ToString());

                        using (Transaction t = new Transaction(linkedDoc, "Highlight Element"))
                        {
                            try
                            {
                                t.Start();

                                linkedView.SetElementOverrides(linkedElement.Id, ogs);
                                //linkedDoc.ActiveView.SetElementOverrides(linkedElementId, ogs);

                                uidoc.RefreshActiveView();
                                t.Commit();

                                return Result.Succeeded;
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Exception", ex.ToString());
                                return Result.Failed;
                            }
                        }
                    }
                    else
                    {
                        //TaskDialog.Show("Error", "Not Happening");
                        return Result.Cancelled;
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void CopyButton_Click(object sender, EventArgs e)
        {
            if (_elementIdLabel != null)
            {
                Clipboard.SetText(_elementIdLabel.Text.Substring("Element ID: ".Length));
            }
        }
    }
}
