﻿using Autodesk.Revit.UI;
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
            application.CreateRibbonTab("Arth");
            string path = Assembly.GetExecutingAssembly().Location;

            RibbonPanel panel1 = application.CreateRibbonPanel("Arth", "Manage Documentation");
            RibbonPanel panel2 = application.CreateRibbonPanel("Arth", "Miscellaneous Tasks");

            PushButtonData button1 = new PushButtonData("Button1", "Add Sheet", path, "MyRevitCommands.AddSheet");

            Uri imagePath1 = new Uri(@"https://www.linkpicture.com/q/p1_26.png");
            BitmapImage image1 = new BitmapImage(imagePath1);

            PushButton pushButton1 = panel1.AddItem(button1) as PushButton;
            pushButton1.LargeImage = image1;

            PushButtonData button2 = new PushButtonData("Button2", "Add Viewport", path, "MyRevitCommands.AddViewport");

            Uri imagePath2 = new Uri(@"https://www.linkpicture.com/q/p2_29.png");
            BitmapImage image2 = new BitmapImage(imagePath2);

            PushButton pushButton2 = panel1.AddItem(button2) as PushButton;
            pushButton2.LargeImage = image2;

            PushButtonData button3 = new PushButtonData("Button3", "Add Tag", path, "MyRevitCommands.AddTag");

            Uri imagePath3 = new Uri(@"https://www.linkpicture.com/q/p4_11.png");
            BitmapImage image3 = new BitmapImage(imagePath3);

            PushButton pushButton3 = panel1.AddItem(button3) as PushButton;
            pushButton3.LargeImage = image3;

            PushButtonData button4 = new PushButtonData("Button4", "Align Tag", path, "MyRevitCommands.AlignTag");

            Uri imagePath4 = new Uri(@"https://www.linkpicture.com/q/p9_2.png");
            BitmapImage image4 = new BitmapImage(imagePath4);

            PushButton pushButton4 = panel1.AddItem(button4) as PushButton;
            pushButton4.LargeImage = image4;

            PushButtonData button5 = new PushButtonData("Button5", "Linked Element Id", path, "MyRevitCommands.GetLinkedElementId");

            Uri imagePath5 = new Uri(@"https://www.linkpicture.com/q/p8_6.png");
            BitmapImage image5 = new BitmapImage(imagePath5);

            PushButton pushButton5 = panel2.AddItem(button5) as PushButton;
            pushButton5.LargeImage = image5;

            PushButtonData button6 = new PushButtonData("Button6", "Export Schedule", path, "MyRevitCommands.Schedule");

            Uri imagePath6 = new Uri(@"https://www.linkpicture.com/q/p5_7.png");
            BitmapImage image6 = new BitmapImage(imagePath6);

            PushButton pushButton6 = panel2.AddItem(button6) as PushButton;
            pushButton6.LargeImage = image6;

            //panel2.AddSeparator();

            return Result.Succeeded;
        }
    }
}