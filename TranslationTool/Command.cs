#region Namespaces

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;

#endregion Namespaces

namespace TranslationTool
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            
            string path = @"W:\S\BIM\Z-LINKED EXCEL\SOM-Chinese Translation.xlsx";
            //used to debug.
            //string path = @"C:\SOM-Chinese Translation.xlsx";

            //Project file path and name to track sheets annotations and titles.
            string centralFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            
            TranslateUpdateDatabase updateDatabase = new TranslateUpdateDatabase();
            updateDatabase.TransactionUpdate(doc, app, path, centralFilePath);

            
            return Result.Succeeded;
        }
    }
}