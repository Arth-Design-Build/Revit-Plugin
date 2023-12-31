﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddPipeTag : IExternalCommand
    {
        private Document _doc;
        private UIDocument _uiDoc;

        /*
        public AddPipeTag(UIApplication uiapp)
        {
            _uiDoc = uiapp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            _app = uiapp.Application;
        }
        */

        private XYZ MoveRight(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X + scaleFactor, point.Y, point.Z);
        }

        private XYZ MoveDown(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X, point.Y - scaleFactor, point.Z);
        }

        private XYZ MoveLeft(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X - scaleFactor, point.Y, point.Z);
        }

        private XYZ MoveUp(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X, point.Y + scaleFactor, point.Z);
        }

        private IEnumerable<XYZ> Shift(int end, XYZ point, double scaleFactor)
        {
            var moves = new List<System.Func<XYZ, double, XYZ>> { MoveRight, MoveDown, MoveLeft, MoveUp };
            var n = 1;
            var pos = point;
            var timesToMove = 1;

            yield return pos;

            while (true)
            {
                for (var i = 0; i < 2; i++)
                {
                    var move = moves[i % 4];
                    for (var j = 0; j < timesToMove; j++)
                    {
                        if (n >= end) yield break;
                        pos = move(pos, scaleFactor);
                        n++;
                        yield return pos;
                    }
                }
                timesToMove++;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document _doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Prompt the user to make a selection
            var selection = uidoc.Selection;
            IList<Element> selectedElements = sel.PickElementsByRectangle();

            var PipeFiltered = new List<Element>();
            foreach (var elem in selectedElements)
            {
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves || elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                {
                    //FamilyInstance Pipe = (FamilyInstance)elem;
                    PipeFiltered.Add(elem);
                }
            }

            var scaleFactor = 5.0;

            var avoidLoc = new List<XYZ>();
            foreach (var d in PipeFiltered)
            {
                BoundingBoxXYZ boundingBox1 = d.get_BoundingBox(_doc.ActiveView);
                XYZ centerPoint = (boundingBox1.Max + boundingBox1.Min) / 2;
                if (centerPoint != null) avoidLoc.Add(new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z));
            }

            // Get a list of available families for Pipes
            FilteredElementCollector fec = new FilteredElementCollector(_doc);
            fec.OfClass(typeof(Family));
            List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Pipe") && f.Name.Contains("Tag")).ToList();
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
            forma.AutoSize = true;
            forma.BackColor = System.Drawing.Color.LightGray;

            GroupBox groupBox = new GroupBox();
            groupBox.Location = new System.Drawing.Point(10, 10);
            groupBox.Size = new Size(480, 140);
            groupBox.AutoSize = true;
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

            System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();
            okButton.Text = "OK";
            okButton.DialogResult = DialogResult.OK;
            okButton.Location = new System.Drawing.Point(180, y+50>160?y+50:160);
            forma.Controls.Add(okButton);

            System.Windows.Forms.Button cancelButton = new System.Windows.Forms.Button();
            cancelButton.Text = "Cancel";
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Location = new System.Drawing.Point(260, y+50>160?y+50:160);
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

            var tagFamily = selectedFamily;

            var tagSymbols = new FilteredElementCollector(_doc)
            .OfCategory(BuiltInCategory.OST_PipeTags)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .FirstOrDefault(fs => fs.Family.Id == tagFamily.Id);

            if (tagSymbols == null)
            {
                TaskDialog.Show("Error", "No Tag Symbols Found");
                return Result.Failed;
            }

            System.Windows.Forms.Form formb = new System.Windows.Forms.Form();
            formb.Text = "Do You Want to Add Leader?";
            formb.StartPosition = FormStartPosition.CenterScreen;
            formb.FormBorderStyle = FormBorderStyle.FixedDialog;
            formb.MinimizeBox = false;
            formb.MaximizeBox = false;
            formb.ShowInTaskbar = false;
            formb.AutoScroll = true;
            formb.ClientSize = new Size(350, 150);
            formb.BackColor = System.Drawing.Color.LightGray;

            GroupBox groupBox1 = new GroupBox();
            groupBox1.Location = new System.Drawing.Point(10, 10);
            groupBox1.Size = new Size(330, 90);
            formb.Controls.Add(groupBox1);

            RadioButton cb1 = new RadioButton();
            cb1.Text = "Yes, Use Leader";
            cb1.AutoSize = true;
            cb1.Location = new System.Drawing.Point(25, 40);
            groupBox1.Controls.Add(cb1);

            RadioButton cb2 = new RadioButton();
            cb2.Text = "No, Don't Use Leader";
            cb2.AutoSize = true;
            cb2.Location = new System.Drawing.Point(165, 40);
            groupBox1.Controls.Add(cb2);

            System.Windows.Forms.Button okButton1 = new System.Windows.Forms.Button();
            okButton1.Text = "OK";
            okButton1.DialogResult = DialogResult.OK;
            okButton1.Location = new System.Drawing.Point(100, 110);
            formb.Controls.Add(okButton1);

            System.Windows.Forms.Button cancelButton1 = new System.Windows.Forms.Button();
            cancelButton1.Text = "Cancel";
            cancelButton1.DialogResult = DialogResult.Cancel;
            cancelButton1.Location = new System.Drawing.Point(180, 110);
            formb.Controls.Add(cancelButton1);

            bool leaderStatus = true;

            if (formb.ShowDialog() == DialogResult.OK)
            {
                foreach (System.Windows.Forms.Control control in groupBox1.Controls)
                {
                    if (control is RadioButton && ((RadioButton)control).Checked)
                    {
                        if (control.Text.ToString() == "Yes, Use Leader")
                        {
                            leaderStatus = true;
                            break;
                        }
                        else
                        {
                            leaderStatus = false;
                            break;
                        }
                    }
                }
            }

            //TaskDialog.Show("Count", PipeFiltered.Count.ToString());

            System.Windows.Forms.Form formc = new System.Windows.Forms.Form();
            formc.Text = "Specify the Minimum Pipe Length in Millimetre";
            formc.StartPosition = FormStartPosition.CenterScreen;
            formc.FormBorderStyle = FormBorderStyle.FixedDialog;
            formc.MinimizeBox = false;
            formc.MaximizeBox = false;
            formc.ShowInTaskbar = false;
            formc.AutoScroll = true;
            formc.ClientSize = new Size(350, 120);
            formc.BackColor = System.Drawing.Color.LightGray;

            GroupBox groupBox2 = new GroupBox();
            groupBox2.Location = new System.Drawing.Point(10, 10);
            groupBox2.Size = new Size(330, 60);
            formc.Controls.Add(groupBox2);

            System.Windows.Forms.TextBox tb = new System.Windows.Forms.TextBox();
            tb.Width = 310;
            tb.Location = new System.Drawing.Point(10, 20);
            groupBox2.Controls.Add(tb);

            System.Windows.Forms.Button okButton2 = new System.Windows.Forms.Button();
            okButton2.Text = "OK";
            okButton2.DialogResult = DialogResult.OK;
            okButton2.Location = new System.Drawing.Point(100, 80);
            formc.Controls.Add(okButton2);

            System.Windows.Forms.Button cancelButton2 = new System.Windows.Forms.Button();
            cancelButton2.Text = "Cancel";
            cancelButton2.DialogResult = DialogResult.Cancel;
            cancelButton2.Location = new System.Drawing.Point(180, 80);
            formc.Controls.Add(cancelButton2);

            string minLength = null;

            if (formc.ShowDialog() == DialogResult.OK)
            {
                foreach (System.Windows.Forms.Control control in groupBox2.Controls)
                {
                    if (control is System.Windows.Forms.TextBox)
                    {
                        minLength = control.Text;
                    }
                }
            }

            //TaskDialog.Show("Minimum Length", minLength);
            double mLength = int.Parse(minLength) * 0.00328084;

            foreach (var d in PipeFiltered)
            {
                BoundingBoxXYZ boundingBox2 = d.get_BoundingBox(_doc.ActiveView);
                XYZ centerPoint = (boundingBox2.Max + boundingBox2.Min) / 2;
                if (centerPoint == null)
                {
                    //TaskDialog.Show("CenterPoint", "Null");
                    continue;
                }

                var maxPoint = boundingBox2.Max;
                var minPoint = boundingBox2.Min;

                var deltaX = maxPoint.X - minPoint.X;
                var deltaY = maxPoint.Y - minPoint.Y;
                var deltaZ = maxPoint.Z - minPoint.Z;

                var distanceXYZ = Math.Sqrt(deltaX * deltaX + deltaY * deltaY + deltaZ * deltaZ);

                if (distanceXYZ < mLength)
                {
                    //TaskDialog.Show("Distance", "Not Staisfying");
                    continue;
                }

                var levelPoint = new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z);
                var R = new Reference(d);
                IndependentTag IT = null;

                using (Transaction tx = new Transaction(_doc, "Tag element"))
                {
                    try
                    {
                        tx.Start();
                        IT = IndependentTag.Create(_doc, uidoc.ActiveView.Id, R, leaderStatus, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, levelPoint);
                        IT.ChangeTypeId(tagSymbols.Id);
                        tx.Commit();
                    }
                    catch
                    {
                        continue;
                    }
                }

                var tagBB = IT.get_BoundingBox((Autodesk.Revit.DB.View)_doc.GetElement(IT.OwnerViewId));
                var globalMax = tagBB.Max;
                var globalMin = tagBB.Min;
                var BBloc = new XYZ((globalMax.X), (globalMax.Y + globalMin.Y) / 2, globalMax.Z);
                var avoidPoints = Shift(50, BBloc, scaleFactor);

                foreach (var point in avoidPoints)
                {
                    var closestPipe = FindClosestPipe(point, avoidLoc);
                    if (closestPipe != null)
                    {
                        var distance = closestPipe.DistanceTo(point);
                        if (distance < scaleFactor)
                        {
                            var direction = (point - closestPipe).Normalize();
                            var newPoint = closestPipe + (direction * scaleFactor);
                            var tx1 = new Transaction(_doc);
                            tx1.Start("Modification of Tags");
                            IT.TagHeadPosition = newPoint;
                            IT.LeaderEndCondition = LeaderEndCondition.Free;
                            tx1.Commit();
                            break;
                        }
                    }
                }
            }
            return Result.Succeeded;
        }

        private XYZ FindClosestPipe(XYZ point, List<XYZ> avoidLoc)
        {
            XYZ closestPipe = null;
            var closestDistance = double.MaxValue;
            foreach (var Pipe in avoidLoc)
            {
                var distance = Pipe.DistanceTo(point);
                if (distance < closestDistance)
                {
                    closestPipe = Pipe;
                    closestDistance = distance;
                }
            }
            return closestPipe;
        }
    }
}
