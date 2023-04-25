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
                Text = "Mechanical Equipments"
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

                    var message1 = "";
                    AddDoorTag addTag = new AddDoorTag();
                    addTag.Execute(commandData, ref message1, elements);
                    return Result.Succeeded;
                }
                else
                {
                    UIApplication uiapp = uidoc.Application;
                    UIDocument uiDoc = uiapp.ActiveUIDocument;
                    Selection sel = uiDoc.Selection;

                    var message1 = "";
                    AddWindowTag addTag = new AddWindowTag();
                    addTag.Execute(commandData, ref message1, elements);
                    return Result.Succeeded;
                }
            }

            else
            {
                if(flag1 == 1)
                {
                    UIApplication uiapp = uidoc.Application;
                    UIDocument uiDoc = uiapp.ActiveUIDocument;
                    Selection sel = uiDoc.Selection;

                    var message1 = "";
                    AddEquipmentTag addTag = new AddEquipmentTag();
                    addTag.Execute(commandData, ref message1, elements);
                    return Result.Succeeded;
                }
                else
                {
                    if(flag3 == 1)
                    {
                        UIApplication uiapp = uidoc.Application;
                        UIDocument uiDoc = uiapp.ActiveUIDocument;
                        Selection sel = uiDoc.Selection;

                        var message1 = "";
                        AddDuctTag addTag = new AddDuctTag();
                        addTag.Execute(commandData, ref message1, elements);
                        return Result.Succeeded;
                    }
                    else
                    {
                        UIApplication uiapp = uidoc.Application;
                        UIDocument uiDoc = uiapp.ActiveUIDocument;
                        Selection sel = uiDoc.Selection;

                        var message1 = "";
                        AddPipeTag addTag = new AddPipeTag();
                        addTag.Execute(commandData, ref message1, elements);
                        return Result.Succeeded;
                    }
                }
            }
        }
    }
}
