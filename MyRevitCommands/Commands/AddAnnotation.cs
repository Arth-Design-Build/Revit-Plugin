using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace MyRevitCommands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class AddAnnotation : Autodesk.Revit.UI.IExternalCommand
    {
        
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            var message1 = "";
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
                Text = "Cable Tray"
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
                Text = "Columns"
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
                Text = "Conducts"
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
                Text = "Ducts"
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
                Text = "Pipes"
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

            var radio5 = new System.Windows.Forms.RadioButton
            {
                Text = "Walls"
            };
            radio5.AutoSize = true;
            radio5.Left = 20;
            radio5.Top = top;
            radio5.CheckedChanged += (sender, args) => {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag6 = 1;
                }
                else
                {
                    flag6 = 0;
                }
            };

            form.Controls.Add(radio5);

            top = top + 25;

            var radio6 = new System.Windows.Forms.RadioButton
            {
                Text = "Grids"
            };
            radio6.AutoSize = true;
            radio6.Left = 20;
            radio6.Top = top;
            radio6.CheckedChanged += (sender, args) =>
            {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag7 = 1;
                }
                else
                {
                    flag7 = 0;
                }
            };

            form.Controls.Add(radio6);

            top = top + 25;

            var radio7 = new System.Windows.Forms.RadioButton
            {
                Text = "Other Imported Components"
            };
            radio7.AutoSize = true;
            radio7.Left = 20;
            radio7.Top = top;
            radio7.CheckedChanged += (sender, args) =>
            {
                var check = (System.Windows.Forms.RadioButton)sender;
                if (check.Checked)
                {
                    flag8 = 1;
                }
                else
                {
                    flag8 = 0;
                }
            };


            //form.Controls.Add(radio7);

            var button = new Button
            {
                Text = "Select",
                Left = 150,
                Top = top + 30,
                AutoSize = true
            };
            form.Controls.Add(button);
            button.Click += (sender, args) =>
            {
                form.Close();
                if(flag1 == 1)
                {

                }
                else
                {
                    if(flag2 == 1)
                    {

                    }
                    else
                    {
                        if(flag3 == 1)
                        {
                            
                        }
                        else
                        {
                            if(flag4 == 1)
                            {
                                AddDimDuct addDimDuct = new AddDimDuct();
                                addDimDuct.Execute(commandData, ref message1, elements);
                            }
                            else
                            {
                                if(flag5 == 1)
                                {
                                    AddDimPipe addDimPipe = new AddDimPipe();
                                    addDimPipe.Execute(commandData, ref message1, elements);
                                }
                                else
                                {
                                    if(flag6 == 1)
                                    {
                                        AddDimWall addDimWall = new AddDimWall();
                                        addDimWall.Execute(commandData, ref message1, elements);
                                    }
                                    else
                                    {
                                        if(flag7 == 1)
                                        {
                                            AddDimGrid addDimGrid = new AddDimGrid();
                                            addDimGrid.Execute(commandData, ref message1, elements);
                                        }
                                        else
                                        {
                                            if(flag8 == 1)
                                            {
                                                AddDimComponent addDimComponent = new AddDimComponent();
                                                addDimComponent.Execute(commandData, ref message1, elements);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            form.ShowDialog();

            return Result.Succeeded;
        }
    }
}