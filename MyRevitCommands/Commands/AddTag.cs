using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            int top = 25;

            int flag1 = 0;
            int flag2 = 0;
            int flag3 = 0;
            int flag4 = 0;
            int flag5 = 0;
            int flag6 = 0;
            int flag7 = 0;
            int flag8 = 0;

            var form = new System.Windows.Forms.Form
            {
                Size = new System.Drawing.Size(400, 400),
                AutoScroll = true
            };
            var radio = new System.Windows.Forms.RadioButton
            {
                Text = "Conduit"
            };
            radio.AutoSize = true;
            radio.Left = 20;
            radio.Top = top;
            radio.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag1 = 1;
                }
                else
                {
                    flag1 = 0;
                }
            };

            form.Controls.Add(radio);

            top = top + 25;

            var radio1 = new System.Windows.Forms.RadioButton
            {
                Text = "Doors"
            };
            radio1.AutoSize = true;
            radio1.Left = 20;
            radio1.Top = top;
            radio1.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag2 = 1;
                }
                else
                {
                    flag2 = 0;
                }
            };

            form.Controls.Add(radio1);

            top = top + 25;

            var radio2 = new System.Windows.Forms.RadioButton
            {
                Text = "Ducts"
            };
            radio2.AutoSize = true;
            radio2.Left = 20;
            radio2.Top = top;
            radio2.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag3 = 1;
                }
                else
                {
                    flag3 = 0;
                }
            };

            form.Controls.Add(radio2);

            top = top + 25;

            var radio3 = new System.Windows.Forms.RadioButton
            {
                Text = "Pipes"
            };
            radio3.AutoSize = true;
            radio3.Left = 20;
            radio3.Top = top;
            radio3.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag4 = 1;
                }
                else
                {
                    flag4 = 0;
                }
            };

            form.Controls.Add(radio3);

            top = top + 25;

            var radio4 = new System.Windows.Forms.RadioButton
            {
                Text = "Windows"
            };
            radio4.AutoSize = true;
            radio4.Left = 20;
            radio4.Top = top;
            radio4.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag5 = 1;
                }
                else
                {
                    flag5 = 0;
                }
            };

            form.Controls.Add(radio4);

            top = top + 25;

            var button1 = new Button
            {
                Text = "Place Tags",
                Left = 140,
                Top = top+20,
            };
            button1.AutoSize = true;
            button1.Click += (sender, args) =>
            {
                form.Close();
            };

            form.Controls.Add(button1);

            form.Text = "Select Elements to Place Tags";
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

            if (flag2 == 1 || flag5 == 1)
            {
                if(flag2 == 1)
                {
                    UIApplication uiapp = uidoc.Application;
                    UIDocument uiDoc = uiapp.ActiveUIDocument;
                    Selection sel = uiDoc.Selection;

                    // Get a list of available families for ducts
                    FilteredElementCollector fec = new FilteredElementCollector(doc);
                    fec.OfClass(typeof(Family));
                    List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Door Tag") || f.Name.Contains("Door_Tag")).ToList();
                    List<string> familyNames = new List<string>();
                    foreach (Family family in families)
                    {
                        familyNames.Add(family.Name);
                    }

                    // Create a form with radio buttons to select the family
                    System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
                    forma.Text = "Select a family";
                    forma.StartPosition = FormStartPosition.CenterScreen;
                    forma.FormBorderStyle = FormBorderStyle.FixedDialog;
                    forma.MinimizeBox = false;
                    forma.MaximizeBox = false;
                    forma.ShowInTaskbar = false;
                    forma.AutoScroll = true;
                    forma.ClientSize = new Size(500, 200);
                    forma.BackColor = System.Drawing.Color.LightGray;

                    GroupBox groupBox = new GroupBox();
                    groupBox.Location = new System.Drawing.Point(10, 10);
                    groupBox.Size = new Size(480, 140);
                    forma.Controls.Add(groupBox);

                    int y = 20;
                    foreach (Family family in families)
                    {
                        RadioButton radioButton = new RadioButton();
                        radioButton.Text = family.Name;
                        radioButton.Location = new System.Drawing.Point(20, y);
                        radioButton.AutoSize = true;
                        radioButton.Tag = family;
                        groupBox.Controls.Add(radioButton);
                        y += 25;
                    }

                    Button okButton = new Button();
                    okButton.Text = "OK";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new System.Drawing.Point(180, 160);
                    forma.Controls.Add(okButton);

                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Location = new System.Drawing.Point(260, 160);
                    forma.Controls.Add(cancelButton);

                    // Show the form and get the selected family
                    Family selectedFamily = null;
                    if (forma.ShowDialog() == DialogResult.OK)
                    {
                        foreach (System.Windows.Forms.Control control in groupBox.Controls)
                        {
                            if (control is RadioButton && ((RadioButton)control).Checked)
                            {
                                selectedFamily = (Family)((RadioButton)control).Tag;
                                break;
                            }
                        }
                    }


                    if (selectedFamily != null)
                    {
                        // Set the tag mode and orientation
                        TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                        TagOrientation tagOrientation = TagOrientation.Horizontal;

                        List<BuiltInCategory> categories = new List<BuiltInCategory>();
                        categories.Add(BuiltInCategory.OST_Doors);

                        ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

                        IList<Element> selectedElements = sel.PickElementsByRectangle();

                        try
                        {
                            // Start a new transaction
                            using (Transaction transaction = new Transaction(doc, "Tag Archi Systems"))
                            {
                                transaction.Start();

                                foreach (Element system in selectedElements)
                                {
                                    if (system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                                    {
                                        Reference reference = new Reference(system);

                                        // Get the center point of the system
                                        BoundingBoxXYZ boundingBox = system.get_BoundingBox(doc.ActiveView);
                                        XYZ centerPoint = (boundingBox.Max + boundingBox.Min) / 2;

                                        // Create the tag
                                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, reference, true, tagMode, tagOrientation, centerPoint);
                                    }
                                }

                                // Commit the transaction
                                transaction.Commit();
                            }

                            return Result.Succeeded;
                        }
                        catch (Exception e)
                        {
                            message = e.Message;
                            return Result.Failed;
                        }
                    }
                    else
                    {
                        return Result.Failed;
                    }
                }
                else
                {
                    UIApplication uiapp = uidoc.Application;
                    UIDocument uiDoc = uiapp.ActiveUIDocument;
                    Selection sel = uiDoc.Selection;

                    // Get a list of available families for ducts
                    FilteredElementCollector fec = new FilteredElementCollector(doc);
                    fec.OfClass(typeof(Family));
                    List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Window Tag") || f.Name.Contains("Window_Tag")).ToList();
                    List<string> familyNames = new List<string>();
                    foreach (Family family in families)
                    {
                        familyNames.Add(family.Name);
                    }

                    // Create a form with radio buttons to select the family
                    System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
                    forma.Text = "Select a family";
                    forma.StartPosition = FormStartPosition.CenterScreen;
                    forma.FormBorderStyle = FormBorderStyle.FixedDialog;
                    forma.MinimizeBox = false;
                    forma.MaximizeBox = false;
                    forma.ShowInTaskbar = false;
                    forma.AutoScroll = true;
                    forma.ClientSize = new Size(500, 200);
                    forma.BackColor = System.Drawing.Color.LightGray;

                    GroupBox groupBox = new GroupBox();
                    groupBox.Location = new System.Drawing.Point(10, 10);
                    groupBox.Size = new Size(480, 140);
                    forma.Controls.Add(groupBox);

                    int y = 20;
                    foreach (Family family in families)
                    {
                        RadioButton radioButton = new RadioButton();
                        radioButton.Text = family.Name;
                        radioButton.Location = new System.Drawing.Point(20, y);
                        radioButton.AutoSize = true;
                        radioButton.Tag = family;
                        groupBox.Controls.Add(radioButton);
                        y += 25;
                    }

                    Button okButton = new Button();
                    okButton.Text = "OK";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new System.Drawing.Point(180, 160);
                    forma.Controls.Add(okButton);

                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Location = new System.Drawing.Point(260, 160);
                    forma.Controls.Add(cancelButton);

                    // Show the form and get the selected family
                    Family selectedFamily = null;
                    if (forma.ShowDialog() == DialogResult.OK)
                    {
                        foreach (System.Windows.Forms.Control control in groupBox.Controls)
                        {
                            if (control is RadioButton && ((RadioButton)control).Checked)
                            {
                                selectedFamily = (Family)((RadioButton)control).Tag;
                                break;
                            }
                        }
                    }


                    if (selectedFamily != null)
                    {
                        // Set the tag mode and orientation
                        TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                        TagOrientation tagOrientation = TagOrientation.Horizontal;

                        List<BuiltInCategory> categories = new List<BuiltInCategory>();
                        categories.Add(BuiltInCategory.OST_Windows);

                        ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

                        IList<Element> selectedElements = sel.PickElementsByRectangle();

                        try
                        {
                            // Start a new transaction
                            using (Transaction transaction = new Transaction(doc, "Tag Archi Systems"))
                            {
                                transaction.Start();

                                foreach (Element system in selectedElements)
                                {
                                    if (system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                                    {
                                        Reference reference = new Reference(system);

                                        // Get the center point of the system
                                        BoundingBoxXYZ boundingBox = system.get_BoundingBox(doc.ActiveView);
                                        XYZ centerPoint = (boundingBox.Max + boundingBox.Min) / 2;

                                        // Create the tag
                                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, reference, true, tagMode, tagOrientation, centerPoint);
                                    }
                                }

                                // Commit the transaction
                                transaction.Commit();
                            }

                            return Result.Succeeded;
                        }
                        catch (Exception e)
                        {
                            message = e.Message;
                            return Result.Failed;
                        }
                    }
                    else
                    {
                        return Result.Failed;
                    }
                }
            }

            else
            {
                if(flag1 == 1)
                {
                    UIApplication uiapp = uidoc.Application;
                    UIDocument uiDoc = uiapp.ActiveUIDocument;
                    Selection sel = uiDoc.Selection;

                    // Get a list of available families for ducts
                    FilteredElementCollector fec = new FilteredElementCollector(doc);
                    fec.OfClass(typeof(Family));
                    List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Conduit Tag") || f.Name.Contains("Conduit_Tag")).ToList();
                    List<string> familyNames = new List<string>();
                    foreach (Family family in families)
                    {
                        familyNames.Add(family.Name);
                    }

                    // Create a form with radio buttons to select the family
                    System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
                    forma.Text = "Select a family";
                    forma.StartPosition = FormStartPosition.CenterScreen;
                    forma.FormBorderStyle = FormBorderStyle.FixedDialog;
                    forma.MinimizeBox = false;
                    forma.MaximizeBox = false;
                    forma.ShowInTaskbar = false;
                    forma.AutoScroll = true;
                    forma.ClientSize = new Size(500, 200);
                    forma.BackColor = System.Drawing.Color.LightGray;

                    GroupBox groupBox = new GroupBox();
                    groupBox.Location = new System.Drawing.Point(10, 10);
                    groupBox.Size = new Size(480, 140);
                    forma.Controls.Add(groupBox);

                    int y = 20;
                    foreach (Family family in families)
                    {
                        RadioButton radioButton = new RadioButton();
                        radioButton.Text = family.Name;
                        radioButton.Location = new System.Drawing.Point(20, y);
                        radioButton.AutoSize = true;
                        radioButton.Tag = family;
                        groupBox.Controls.Add(radioButton);
                        y += 25;
                    }

                    Button okButton = new Button();
                    okButton.Text = "OK";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new System.Drawing.Point(180, 160);
                    forma.Controls.Add(okButton);

                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Location = new System.Drawing.Point(260, 160);
                    forma.Controls.Add(cancelButton);

                    // Show the form and get the selected family
                    Family selectedFamily = null;
                    if (forma.ShowDialog() == DialogResult.OK)
                    {
                        foreach (System.Windows.Forms.Control control in groupBox.Controls)
                        {
                            if (control is RadioButton && ((RadioButton)control).Checked)
                            {
                                selectedFamily = (Family)((RadioButton)control).Tag;
                                break;
                            }
                        }
                    }


                    if (selectedFamily != null)
                    {
                        // Set the tag mode and orientation
                        TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                        TagOrientation tagOrientation = TagOrientation.Horizontal;

                        // Create a list of categories for MEP systems
                        List<BuiltInCategory> categories = new List<BuiltInCategory>();
                        categories.Add(BuiltInCategory.OST_Conduit);
                        categories.Add(BuiltInCategory.OST_ConduitFitting);

                        // Create a filter for MEP systems
                        ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

                        // Get a list of MEP systems in the active view
                        IList<Element> selectedElements = sel.PickElementsByRectangle();

                        try
                        {
                            // Start a new transaction
                            using (Transaction transaction = new Transaction(doc, "Tag MEP Systems"))
                            {
                                transaction.Start();

                                // Loop through each MEP system and create a tag
                                foreach (Element system in selectedElements)
                                {
                                    if (system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit || system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_ConduitFitting)
                                    {
                                        Reference reference = new Reference(system);

                                        // Get the center point of the system
                                        BoundingBoxXYZ boundingBox = system.get_BoundingBox(doc.ActiveView);
                                        XYZ centerPoint = (boundingBox.Max + boundingBox.Min) / 2;

                                        // Create the tag
                                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, reference, true, tagMode, tagOrientation, centerPoint);
                                    }
                                }

                                // Commit the transaction
                                transaction.Commit();
                            }

                            return Result.Succeeded;
                        }
                        catch (Exception e)
                        {
                            message = e.Message;
                            return Result.Failed;
                        }
                    }
                    else
                    {
                        return Result.Failed;
                    }
                }
                else
                {
                    if(flag3 == 1)
                    {
                        UIApplication uiapp = uidoc.Application;
                        UIDocument uiDoc = uiapp.ActiveUIDocument;
                        Selection sel = uiDoc.Selection;

                        // Get a list of available families for ducts
                        FilteredElementCollector fec = new FilteredElementCollector(doc);
                        fec.OfClass(typeof(Family));
                        List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Duct Tag") || f.Name.Contains("Duct_Tag")).ToList();
                        List<string> familyNames = new List<string>();
                        foreach (Family family in families)
                        {
                            familyNames.Add(family.Name);
                        }

                        // Create a form with radio buttons to select the family
                        System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
                        forma.Text = "Select a family";
                        forma.StartPosition = FormStartPosition.CenterScreen;
                        forma.FormBorderStyle = FormBorderStyle.FixedDialog;
                        forma.MinimizeBox = false;
                        forma.MaximizeBox = false;
                        forma.ShowInTaskbar = false;
                        forma.AutoScroll = true;
                        forma.ClientSize = new Size(500, 200);
                        forma.BackColor = System.Drawing.Color.LightGray;

                        GroupBox groupBox = new GroupBox();
                        groupBox.Location = new System.Drawing.Point(10, 10);
                        groupBox.Size = new Size(480, 140);
                        forma.Controls.Add(groupBox);

                        int y = 20;
                        foreach (Family family in families)
                        {
                            RadioButton radioButton = new RadioButton();
                            radioButton.Text = family.Name;
                            radioButton.Location = new System.Drawing.Point(20, y);
                            radioButton.AutoSize = true;
                            radioButton.Tag = family;
                            groupBox.Controls.Add(radioButton);
                            y += 25;
                        }

                        Button okButton = new Button();
                        okButton.Text = "OK";
                        okButton.DialogResult = DialogResult.OK;
                        okButton.Location = new System.Drawing.Point(180, 160);
                        forma.Controls.Add(okButton);

                        Button cancelButton = new Button();
                        cancelButton.Text = "Cancel";
                        cancelButton.DialogResult = DialogResult.Cancel;
                        cancelButton.Location = new System.Drawing.Point(260, 160);
                        forma.Controls.Add(cancelButton);

                        // Show the form and get the selected family
                        Family selectedFamily = null;
                        if (forma.ShowDialog() == DialogResult.OK)
                        {
                            foreach (System.Windows.Forms.Control control in groupBox.Controls)
                            {
                                if (control is RadioButton && ((RadioButton)control).Checked)
                                {
                                    selectedFamily = (Family)((RadioButton)control).Tag;
                                    break;
                                }
                            }
                        }


                        if (selectedFamily != null)
                        {
                            // Set the tag mode and orientation
                            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                            TagOrientation tagOrientation = TagOrientation.Horizontal;

                            // Create a list of categories for MEP systems
                            List<BuiltInCategory> categories = new List<BuiltInCategory>();
                            categories.Add(BuiltInCategory.OST_DuctCurves);
                            categories.Add(BuiltInCategory.OST_DuctFitting);

                            // Create a filter for MEP systems
                            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

                            // Get a list of MEP systems in the active view
                            IList<Element> selectedElements = sel.PickElementsByRectangle();

                            try
                            {
                                // Start a new transaction
                                using (Transaction transaction = new Transaction(doc, "Tag MEP Systems"))
                                {
                                    transaction.Start();

                                    // Loop through each MEP system and create a tag
                                    foreach (Element system in selectedElements)
                                    {
                                        if (system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves || system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctFitting)
                                        {
                                            Reference reference = new Reference(system);

                                            // Get the center point of the system
                                            BoundingBoxXYZ boundingBox = system.get_BoundingBox(doc.ActiveView);
                                            XYZ centerPoint = (boundingBox.Max + boundingBox.Min) / 2;

                                            // Create the tag
                                            IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, reference, true, tagMode, tagOrientation, centerPoint);
                                        }
                                    }

                                    // Commit the transaction
                                    transaction.Commit();
                                }

                                return Result.Succeeded;
                            }
                            catch (Exception e)
                            {
                                message = e.Message;
                                return Result.Failed;
                            }
                        }
                        else
                        {
                            return Result.Failed;
                        }
                    }
                    else
                    {
                        UIApplication uiapp = uidoc.Application;
                        UIDocument uiDoc = uiapp.ActiveUIDocument;
                        Selection sel = uiDoc.Selection;

                        // Get a list of available families for ducts
                        FilteredElementCollector fec = new FilteredElementCollector(doc);
                        fec.OfClass(typeof(Family));
                        List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Pipe Tag") || f.Name.Contains("Pipe_Tag")).ToList();
                        List<string> familyNames = new List<string>();
                        foreach (Family family in families)
                        {
                            familyNames.Add(family.Name);
                        }

                        // Create a form with radio buttons to select the family
                        System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
                        forma.Text = "Select a family";
                        forma.StartPosition = FormStartPosition.CenterScreen;
                        forma.FormBorderStyle = FormBorderStyle.FixedDialog;
                        forma.MinimizeBox = false;
                        forma.MaximizeBox = false;
                        forma.ShowInTaskbar = false;
                        forma.AutoScroll = true;
                        forma.ClientSize = new Size(500, 200);
                        forma.BackColor = System.Drawing.Color.LightGray;

                        GroupBox groupBox = new GroupBox();
                        groupBox.Location = new System.Drawing.Point(10, 10);
                        groupBox.Size = new Size(480, 140);
                        forma.Controls.Add(groupBox);

                        int y = 20;
                        foreach (Family family in families)
                        {
                            RadioButton radioButton = new RadioButton();
                            radioButton.Text = family.Name;
                            radioButton.Location = new System.Drawing.Point(20, y);
                            radioButton.AutoSize = true;
                            radioButton.Tag = family;
                            groupBox.Controls.Add(radioButton);
                            y += 25;
                        }

                        Button okButton = new Button();
                        okButton.Text = "OK";
                        okButton.DialogResult = DialogResult.OK;
                        okButton.Location = new System.Drawing.Point(180, 160);
                        forma.Controls.Add(okButton);

                        Button cancelButton = new Button();
                        cancelButton.Text = "Cancel";
                        cancelButton.DialogResult = DialogResult.Cancel;
                        cancelButton.Location = new System.Drawing.Point(260, 160);
                        forma.Controls.Add(cancelButton);

                        // Show the form and get the selected family
                        Family selectedFamily = null;
                        if (forma.ShowDialog() == DialogResult.OK)
                        {
                            foreach (System.Windows.Forms.Control control in groupBox.Controls)
                            {
                                if (control is RadioButton && ((RadioButton)control).Checked)
                                {
                                    selectedFamily = (Family)((RadioButton)control).Tag;
                                    break;
                                }
                            }
                        }


                        if (selectedFamily != null)
                        {
                            // Set the tag mode and orientation
                            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                            TagOrientation tagOrientation = TagOrientation.Horizontal;

                            // Create a list of categories for MEP systems
                            List<BuiltInCategory> categories = new List<BuiltInCategory>();
                            categories.Add(BuiltInCategory.OST_PipeCurves);
                            categories.Add(BuiltInCategory.OST_PipeFitting);

                            // Create a filter for MEP systems
                            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

                            // Get a list of MEP systems in the active view
                            IList<Element> selectedElements = sel.PickElementsByRectangle();

                            try
                            {
                                // Start a new transaction
                                using (Transaction transaction = new Transaction(doc, "Tag MEP Systems"))
                                {
                                    transaction.Start();

                                    // Loop through each MEP system and create a tag
                                    foreach (Element system in selectedElements)
                                    {
                                        if (system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves || system.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                                        {
                                            Reference reference = new Reference(system);

                                            // Get the center point of the system
                                            BoundingBoxXYZ boundingBox = system.get_BoundingBox(doc.ActiveView);
                                            XYZ centerPoint = (boundingBox.Max + boundingBox.Min) / 2;

                                            // Create the tag
                                            IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, reference, true, tagMode, tagOrientation, centerPoint);
                                        }
                                    }

                                    // Commit the transaction
                                    transaction.Commit();
                                }

                                return Result.Succeeded;
                            }
                            catch (Exception e)
                            {
                                message = e.Message;
                                return Result.Failed;
                            }
                        }
                        else
                        {
                            return Result.Failed;
                        }
                    }
                }
            }
        }
    }
}
