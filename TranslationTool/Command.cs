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
            string path = @"W:\S\BIM\Z-LINKED EXCEL\SOM-Chinese Translation.xlsx";
            //used to debug.
            //string path = @"C:\SOM-Chinese Translation.xlsx";

            //Project file path and name to track sheets annotations and titles.
            string centralFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            //Translation for Annotations, Title on Sheets and Sheet Names
            translation.StartTranslation(doc, app, path, centralFilePath);
            // Annotations which are not translated. 
            List<string[]> UpdateExcelTranslationGenericAnno = translation._UpdateExcelTranslationGenericAnno;
            Dictionary<string, string[]> Anno_IDandEnglishDictionary = new Dictionary<string,string[]>();

            // Title on sheets not translated
            List<string[]> NotTranslatedTitleOnSheet = translation._NotTranslatedTitleOnSheet;
            // Title on sheet which was on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTitleOnSheet = translation._HoldIsFoundInRevitTitleOnSheet;
            // Dictionary of Key and Values from Excel
            Dictionary<string, string[]> TitleOnSheet_IDEnglishChineseDict = translation._TitleOnSheet_IDEnglishChineseDict;
            // Sheet names not translated
            List<string[]> NotTranslatedSheets = translation._NotTranslatedSheets;
            // Sheet which is on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTSheets = translation._HoldIsFoundInRevitSheets;
            // Dictionary of Key and Values from Excel
            Dictionary<string, string[]> Sheet_IDEnglishChineseDict = translation._Sheet_IDEnglishChineseDict;
            string BuildingName = doc.ProjectInformation.BuildingName;

            // Get project paramter Author number to assign color. 
            var BuildingColor = doc.ProjectInformation.Author;

            //Annotation Excel Sheet
            if (UpdateExcelTranslationGenericAnno.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(Anno_IDandEnglishDictionary, UpdateExcelTranslationGenericAnno, path, 1, "");
            }
            //Title on Sheet Excel Sheet
            if (NotTranslatedTitleOnSheet.Count > 0)
            {
                // combine dictionarys to update Excel file. 
                HoldIsFoundInRevitTitleOnSheet.AddRange(NotTranslatedTitleOnSheet);
            }
            //Title on Sheet Hold was found
            if (HoldIsFoundInRevitTitleOnSheet.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(TitleOnSheet_IDEnglishChineseDict, HoldIsFoundInRevitTitleOnSheet, path, 2, BuildingColor);
            }
            //Sheet Name Excel Sheet
            if (NotTranslatedSheets.Count > 0)
            {
                // combine dictionarys to update Excel file. 
                HoldIsFoundInRevitTSheets.AddRange(NotTranslatedSheets);
            }

            //Sheet Hold was found
            if (HoldIsFoundInRevitTSheets.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(Sheet_IDEnglishChineseDict, HoldIsFoundInRevitTSheets, path, 3, BuildingColor);
            }
            return Result.Succeeded;
        }
    }
}