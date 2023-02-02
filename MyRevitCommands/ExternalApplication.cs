using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace MyRevitCommands
{
    internal class ExternalApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            application.CreateRibbonTab("Arth Design");
            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData button = new PushButtonData("Button1", "Export Schedules", path, "MyRevitCommands.Schedule");
            RibbonPanel panel = application.CreateRibbonPanel("Arth Design", "Batch Export");

            Uri imagePath = new Uri(@"https://www.linkpicture.com/q/batch_export.png");
            BitmapImage image = new BitmapImage(imagePath);

            PushButton pushButton = panel.AddItem(button) as PushButton;
            pushButton.LargeImage = image;

            PushButtonData button1 = new PushButtonData("Button2", "Import Sheet Information", path, "MyRevitCommands.PlaceViewport");
            RibbonPanel panel1 = application.CreateRibbonPanel("Arth Design", "Generate Sheets");

            Uri imagePath1 = new Uri(@"https://www.linkpicture.com/q/2085465-removebg-preview.png");
            BitmapImage image1 = new BitmapImage(imagePath1);

            PushButton pushButton1 = panel1.AddItem(button1) as PushButton;
            pushButton1.LargeImage = image1;

            return Result.Succeeded;
        }
    }
}