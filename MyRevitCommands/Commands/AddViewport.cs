using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class AddViewport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the Revit application and document
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            int flag = 0;

            // Get all sheets in the project
            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc);
            ICollection<Element> sheets = sheetCollector.OfClass(typeof(ViewSheet)).ToElements();

            // Create a dictionary to hold the placed views for each sheet
            Dictionary<ViewSheet, List<Autodesk.Revit.DB.View>> placedViews = new Dictionary<ViewSheet, List<Autodesk.Revit.DB.View>>();

            // Loop through each sheet and get the views placed on it
            foreach (ViewSheet sheet in sheets)
            {
                ICollection<ElementId> viewIds = sheet.GetAllPlacedViews();
                List<Autodesk.Revit.DB.View> views1 = new List<Autodesk.Revit.DB.View>();
                foreach (ElementId viewId in viewIds)
                {
                    Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                    views1.Add(view);
                }
                placedViews[sheet] = views1;
            }

            // Get all views in the project
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
            ICollection<Element> views = viewCollector.OfClass(typeof(Autodesk.Revit.DB.View))
                .WhereElementIsNotElementType()
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => (v.ViewType == ViewType.FloorPlan ||
                             v.ViewType == ViewType.Elevation ||
                             v.ViewType == ViewType.Section ||
                             v.ViewType == ViewType.Rendering ||
                             v.ViewType == ViewType.ThreeD ||
                             v.ViewType == ViewType.DraftingView))
                .Cast<Element>()
                .ToList();

            // Get all views placed on sheets
            IEnumerable<ElementId> placedViewIds = placedViews.Values.SelectMany(v => v).Select(v => v.Id);

            // Get all unplaced views
            IEnumerable<Element> unplacedViews = views.Where(v => !placedViewIds.Contains(v.Id));

            // Create a list to hold the unplaced view names
            List<string> unplacedViewNames = unplacedViews.Select(v => v.Name).ToList();
            List<string> unplacedViewNames1 = unplacedViews.Select(v => v.ToString().Split('.')[3] + ":" + v.Name).ToList();

            // Create a list to hold the sheet and view data
            List<object[]> sheetAndViewData = new List<object[]>();
            foreach (ViewSheet sheet in sheets)
            {
                // Get the sheet number and name
                string sheetNumber = sheet.SheetNumber;
                string sheetName = sheet.Name;

                // Get the placed views
                List<string> placedViewNames = new List<string>();
                foreach (Autodesk.Revit.DB.View view in placedViews[sheet])
                {
                    placedViewNames.Add(view.ToString().Split('.')[3] + ":" + view.Name);
                }

                // Create a checked list box to hold the unplaced view names
                CheckedListBox checkedListBox = new CheckedListBox();
                checkedListBox.CheckOnClick = true;
                checkedListBox.Width = 200;

                placedViewNames.Sort();
                unplacedViewNames1.Sort();

                checkedListBox.Items.AddRange(unplacedViewNames1.ToArray());
                // Create a button to add the selected views to the sheet
                Button addViewsButton = new Button();
                addViewsButton.Text = "Add Views";
                addViewsButton.Width = 100;
                addViewsButton.Click += new EventHandler(delegate (object sender, EventArgs e)
                {
                // Get the selected views
                List<Autodesk.Revit.DB.View> selectedViews = new List<Autodesk.Revit.DB.View>();
                foreach (string selectedViewName in checkedListBox.CheckedItems)
                {
                    Autodesk.Revit.DB.View selectedView = null;
                    foreach (Autodesk.Revit.DB.View view in views)
                    {
                        if (view.Name == selectedViewName.Split(':')[1])
                        {
                            selectedView = view;
                            break;
                        }
                    }
                    selectedViews.Add(selectedView);
                }

                    using (Transaction transaction = new Transaction(doc, "Place Views"))
                    {
                        transaction.Start();

                        int count = 0;
                        BoundingBoxUV outline = sheet.Outline;

                        // calculate the maximum width and height of a viewport based on the sheet outline
                        double maxViewportWidth = outline.Max.U - outline.Min.U;
                        double maxViewportHeight = outline.Max.V - outline.Min.V;

                        // calculate the initial placement point at the top left of the sheet
                        double x = outline.Min.U + maxViewportWidth / 3.75;
                        double y = outline.Max.V - maxViewportHeight / 3.75;

                        // define the spacing between viewports
                        double viewportSpacing = 0.025 * maxViewportWidth;

                        foreach (Autodesk.Revit.DB.View selectedView in selectedViews)
                        {
                            // get the viewport width
                            double viewportWidth = selectedView.Outline.Max.U - selectedView.Outline.Min.U;

                            // check if the viewport is going to exceed the width of the sheet
                            if (x + viewportWidth / 2 > outline.Max.U)
                            {
                                // move to the next line
                                x = outline.Min.U + maxViewportWidth / 4;
                                y -= maxViewportHeight * 0.5;
                            }

                            // create the viewport at the calculated position
                            XYZ point = new XYZ(x, y, 0);
                            Viewport viewport = Viewport.Create(doc, sheet.Id, selectedView.Id, point);

                            // increment the x position for the next viewport
                            x += viewportWidth / 2 + (viewportSpacing*12);

                            count++;
                        }

                        transaction.Commit();
                        //MessageBox.Show("Views Placed Successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Task.Delay(1000).Wait();
                    }

                });

                // Add the sheet and view data to the list
                object[] data = new object[] { sheetNumber, sheetName, placedViewNames.ToArray(), checkedListBox, addViewsButton };
                sheetAndViewData.Add(data);
            }

            // Create a form to display the sheet and view data
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            form.Text = "Add Viewport to Sheet";
            //form.Width = 1000;
            //form.Height = 500;
            form.Size = new System.Drawing.Size(1000, 500);
            form.Padding = new System.Windows.Forms.Padding(10,10,0,10);

            // Create a table layout panel to organize the sheet and view data
            TableLayoutPanel table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.ColumnCount = 5;
            table.AutoScroll = true;
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));

            // Add headers to the table
            Label sheetNumberHeader = new Label();
            sheetNumberHeader.Text = "Sheet No.";
            sheetNumberHeader.Dock = DockStyle.Fill;
            sheetNumberHeader.Font = new System.Drawing.Font(sheetNumberHeader.Font, System.Drawing.FontStyle.Bold);
            table.Controls.Add(sheetNumberHeader, 0, 0);

            Label sheetNameHeader = new Label();
            sheetNameHeader.Text = "Sheet Name";
            sheetNameHeader.Dock = DockStyle.Fill;
            sheetNameHeader.Font = new System.Drawing.Font(sheetNameHeader.Font, System.Drawing.FontStyle.Bold);
            table.Controls.Add(sheetNameHeader, 1, 0);

            Label placedViewHeader = new Label();
            placedViewHeader.Text = "Placed Views";
            placedViewHeader.Dock = DockStyle.Fill;
            placedViewHeader.Font = new System.Drawing.Font(placedViewHeader.Font, System.Drawing.FontStyle.Bold);
            table.Controls.Add(placedViewHeader, 2, 0);

            Label unplacedViewHeader = new Label();
            unplacedViewHeader.Text = "Unplaced Views";
            unplacedViewHeader.Dock = DockStyle.Fill;
            unplacedViewHeader.Font = new System.Drawing.Font(unplacedViewHeader.Font, System.Drawing.FontStyle.Bold);
            table.Controls.Add(unplacedViewHeader, 3, 0);

            Label addButtonHeader = new Label();
            addButtonHeader.Text = "Add";
            addButtonHeader.Dock = DockStyle.Fill;
            addButtonHeader.Font = new System.Drawing.Font(addButtonHeader.Font, System.Drawing.FontStyle.Bold);
            table.Controls.Add(addButtonHeader, 4, 0);

            // Add the table to the form
            form.Controls.Add(table);

            // Add the sheet and view data to the table
            int row = 1;
            foreach (object[] data in sheetAndViewData)
            {
                Label sheetNumberLabel = new Label();
                sheetNumberLabel.Text = data[0].ToString();
                sheetNumberLabel.Dock = DockStyle.Fill;
                table.Controls.Add(sheetNumberLabel, 0, row);

                Label sheetNameLabel = new Label();
                sheetNameLabel.Text = data[1].ToString();
                sheetNameLabel.Dock = DockStyle.Fill;
                table.Controls.Add(sheetNameLabel, 1, row);

                Label placedViewsLabel = new Label();
                placedViewsLabel.Text = string.Join(" ⎥ ", (data[2] as string[]).Select(s => $"{s}"));
                placedViewsLabel.Dock = DockStyle.Fill;
                table.Controls.Add(placedViewsLabel, 2, row);

                System.Windows.Forms.CheckedListBox comboBox = data[3] as System.Windows.Forms.CheckedListBox;
                //comboBox.Dock = DockStyle.Fill;
                table.Controls.Add(comboBox, 3, row);
         
                Button addButton = data[4] as Button;
                //addButton.Dock = DockStyle.Fill;
                table.Controls.Add(addButton, 4, row);        

                row++;
            }

            if (flag == 1)
            {
                form.Close();
            }

            // Define the border style of the form to a dialog box.
            form.FormBorderStyle = FormBorderStyle.FixedDialog;

            // Set the MaximizeBox to false to remove the maximize box.
            form.MaximizeBox = false;

            // Set the MinimizeBox to false to remove the minimize box.
            form.MinimizeBox = false;

            form.Icon = new Icon("favicon.ico");

            form.BackColor = System.Drawing.Color.LightGray;

            form.ShowDialog();

            return Result.Succeeded;
        }
    }
}