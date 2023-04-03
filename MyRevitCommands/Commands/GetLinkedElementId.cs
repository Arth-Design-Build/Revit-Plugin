using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows.Forms;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class GetLinkedElementId : IExternalCommand
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
                Document doc = uidoc.Document;

                // Get the selection from the user
                Reference selectedRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.LinkedElement, "Select an element in a linked model");
                if (selectedRef == null)
                {
                    return Result.Cancelled;
                }

                // Get the linked document and element ID
                RevitLinkInstance linkInstance = doc.GetElement(selectedRef) as RevitLinkInstance;
                if (linkInstance == null)
                {
                    return Result.Failed;
                }
                Document linkedDoc = linkInstance.GetLinkDocument();
                ElementId linkedElementId = selectedRef.LinkedElementId;

                // Get the linked element and display its ID in a form
                Element linkedElement = linkedDoc.GetElement(linkedElementId);
                _form = new System.Windows.Forms.Form();
                _form.Width = 200;
                _form.Height = 200;
                _form.Text = "Element ID of Linked Model";
                _elementIdLabel = new Label();
                _elementIdLabel.Text = "Element ID: " + linkedElement.Id.IntegerValue.ToString();
                _elementIdLabel.AutoSize = true;
                _elementIdLabel.Location = new System.Drawing.Point(15, 25);
                _copyButton = new Button();
                _copyButton.AutoSize = true;
                _copyButton.Text = "Copy";
                _copyButton.Location = new System.Drawing.Point(55, 70);
                _copyButton.Click += new EventHandler(CopyButton_Click);
                _form.Controls.Add(_elementIdLabel);
                _form.Controls.Add(_copyButton);
                _form.StartPosition = FormStartPosition.CenterScreen;
                _form.FormBorderStyle = FormBorderStyle.FixedDialog;
                _form.MaximizeBox = false;
                _form.MinimizeBox = false;
                _form.ShowDialog();

                return Result.Succeeded;
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
