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

            Translate translation = new Translate();
            //string path = @"W:\S\BIM\Z-LINKED EXCEL\SOM-Chinese Translation.xlsx";
            //used to debug.
            string path = @"C:\SOM-Chinese Translation.xlsx";

            //Project file path and name to track sheets annotations and titles.
            string centralFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            //Translation for Annotations, Title on Sheets and Sheet Names
            translation.StartTranslation(doc, app, path, centralFilePath);
            // Annotations which are not translated. 
            List<string[]> UpdateExcelTranslationGenericAnno = translation._UpdateExcelTranslationGenericAnno;
            // Title on sheets not translated
            List<string[]> NotTranslatedTitleOnSheet = translation._NotTranslatedTitleOnSheet;
            // Title on sheet which was on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTitleOnSheet = translation._HoldIsFoundInRevitTitleOnSheet;
            // Sheet names not translated
            List<string[]> NotTranslatedSheets = translation._NotTranslatedSheets;
            // Sheet which is on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTSheets = translation._HoldIsFoundInRevitSheets;

            //string BuildingName = doc.ProjectInformation.BuildingName;
            //TODO Check parameter for Color and Chinses and English Translation if none exist add 
            // parameter. 
            string BuildingColor = doc.ProjectInformation.Author;

            //Annotation Excel Sheet
            if (UpdateExcelTranslationGenericAnno.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(UpdateExcelTranslationGenericAnno, path, 1, "");
            }
            //Title on Sheet Excel Sheet
            if (NotTranslatedTitleOnSheet.Count > 0)
            {
                //ExcelSheet Excel = new ExcelSheet();
                //Excel.Update(NotTranslatedTitleOnSheet, path, 2, BuildingColor);
                HoldIsFoundInRevitTitleOnSheet.AddRange(NotTranslatedTitleOnSheet);
            }
            //Title on Sheet Hold was found
            if (HoldIsFoundInRevitTitleOnSheet.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(HoldIsFoundInRevitTitleOnSheet, path, 2, BuildingColor);
            }
            //Sheet Name Excel Sheet
            if (NotTranslatedSheets.Count > 0)
            {
                //ExcelSheet Excel = new ExcelSheet();
                //Excel.Update(NotTranslatedSheets, path, 3, BuildingColor);
                HoldIsFoundInRevitTSheets.AddRange(NotTranslatedSheets);
            }

            //Sheet Hold was found
            if (HoldIsFoundInRevitTSheets.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(HoldIsFoundInRevitTSheets, path, 3, BuildingColor);
            }
            return Result.Succeeded;
        }
    }
}