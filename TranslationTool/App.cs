#region Namespaces
using System;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Media.Imaging;
#endregion

namespace TranslationTool
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            RibbonPanel alignViewsPanel = ribbonPanel(a);

            //Add button to the panel
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButton alignViewportToCell_Button = alignViewsPanel.AddItem(new PushButtonData
                ("Translation", "Translation", thisAssemblyPath, "TranslationTool.Command")) as PushButton;

            //Tooltip 
            alignViewportToCell_Button.ToolTip = "Translate between English and Chinese";

            ////Set globel directory
            var globePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CH2EN.png");

            //Large image 
            Uri uriImage = new Uri(globePath);
            BitmapImage largeimage = new BitmapImage(uriImage);
            alignViewportToCell_Button.LargeImage = largeimage;

            a.ApplicationClosing += a_ApplicationClosing;

            //Set Application to Idling
            a.Idling += a_Idling;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        public RibbonPanel ribbonPanel(UIControlledApplication a)
        {
            RibbonPanel ribbonPanel = null;

            //Create add-in to the SOM tool ribbon tab
            try
            {
                a.CreateRibbonTab("SOM Tools");
            }
            catch (Exception)
            { }
            //Create Ribbon Panel 
            try
            {
                RibbonPanel alignViewsPanel = a.CreateRibbonPanel("SOM Tools", "Translation");

            }

            catch (Exception)
            { }

            List<RibbonPanel> alignViewpanels = a.GetRibbonPanels("SOM Tools");
            foreach (RibbonPanel panel in alignViewpanels)
            {
                if (panel.Name == "Translation")
                {
                    ribbonPanel = panel;
                }
            }
            return ribbonPanel;
        }

        void a_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {

        }

        void a_ApplicationClosing(object sender, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
