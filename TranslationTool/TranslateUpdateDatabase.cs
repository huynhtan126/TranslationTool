using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;


namespace TranslationTool
{
    class TranslateUpdateDatabase
    {

        public void TransactionUpdate(Document doc, Application app, string path, string centralFilePath)
        {
            #region READ DATABASE
            //Get data from the database to update. 
            TranslateReadDatabase readDatabase = new TranslateReadDatabase();
            //Translation for Annotations, Title on Sheets and Sheet Names
            readDatabase.ReadDatabase(doc, app, path, centralFilePath);

            //Get Annotation data from database
            Dictionary<string, string[]> AnnotationDatabase = readDatabase._AnnotationDatabase;
            // Annotations which are not translated.
            List<string[]> UpdateDatabase_GenericAnnotationy = readDatabase._UpdateDatabase_GenericAnnotation;
            
            // Title on sheets not translated
            List<string[]> NotTranslatedTitleOnSheet = readDatabase._NotTranslatedTitleOnSheet;
            // Title on sheet which was on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTitleOnSheet = readDatabase._HoldIsFoundInRevitTitleOnSheet;
            // Dictionary of Key and Values from Excel
            Dictionary<string, string[]> TitleOnSheet_IDEnglishChineseDict = readDatabase._TitleOnSheet_IDEnglishChineseDict;

            // Sheet names not translated
            List<string[]> NotTranslatedSheets = readDatabase._NotTranslatedSheets;
            // Sheet which is on hold is found in Revit.
            List<string[]> HoldIsFoundInRevitTSheets = readDatabase._HoldIsFoundInRevitSheets;
            // Dictionary of Key and Values from Excel
            Dictionary<string, string[]> Sheet_IDEnglishChineseDict = readDatabase._Sheet_IDEnglishChineseDict;
            string BuildingName = doc.ProjectInformation.BuildingName;
            #endregion

            #region UPDATE DATABASE
            // Get project paramter Author number to assign color. 
            var BuildingColor = doc.ProjectInformation.Author;

            //Annotation Excel Sheet
            if (UpdateDatabase_GenericAnnotationy.Count > 0)
            {
                ExcelSheet Excel = new ExcelSheet();
                Excel.Update(AnnotationDatabase, UpdateDatabase_GenericAnnotationy, path, 1, "");
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
            #endregion
        }
    }
}
