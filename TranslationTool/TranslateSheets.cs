using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace TranslationTool
{
    internal class TranslateSheets
    {
        public List<string[]> _HoldIsFoundInRevitSheets { get; set; }

        //***********************************SheetTranslation***********************************
        public List<string[]> SheetTranslation(Document doc,
            Dictionary<string, string[]> Sheet_IDEnglishChineseDict,
            Dictionary<string, string[]> Sheet_HoldDict,
            List<string> SheetNameCompareList,
            string path, string centralFilePath)
        {
            ExcelSheet Excel = new ExcelSheet();
            RevitModelElements revit = new RevitModelElements();

            List<ViewSheet> sheets = revit.GetAllSheets(doc);
            //Filter to check only the sheets in current docuement.
            Dictionary<string, string[]> SheetsToDeleteFromExcel = RemoveItemFromExcel(Sheet_IDEnglishChineseDict, centralFilePath);

            List<string[]> NotTranslated = new List<string[]>();
            List<string[]> HoldIsFoundInRevit = new List<string[]>();

            foreach (Element e in sheets)
            {
                Parameter Chinese_param = e.LookupParameter("Sheet Name (Alternate Language)");
                Parameter English_param = e.LookupParameter("Sheet Name");
                string English = revit.GetParameterValue(English_param);
                string Chinese = revit.GetParameterValue(Chinese_param);

                ViewSheet sheet = e as ViewSheet;
                Parameter sheet_param = sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER);
                string sheetNumber = revit.GetParameterValue(sheet_param);

                //Check if the sheet is a Z sheet.
                if (sheetNumber.StartsWith("z") || sheetNumber.StartsWith("Z"))
                    continue;

                if (Sheet_IDEnglishChineseDict.ContainsKey(sheetNumber))
                {
                    if (Chinese == "" || Chinese == null)
                    {
                        string[] Values = Sheet_IDEnglishChineseDict[sheetNumber];
                        // Set English note.
                        Transaction t = new Transaction(doc, "Translation by ID");
                        t.Start();
                        // Set english parameter.
                        English_param.Set(Values[0]);
                        Chinese_param.Set(Values[1]);
                        t.Commit();
                    }
                    //If view exists in Revit remove from list.
                    SheetsToDeleteFromExcel.Remove(sheetNumber);
                }

                //Check if Project missing, but translation exists
                //in Revit needs to be translated and project added to database.
                if (Sheet_HoldDict.ContainsKey(sheetNumber))
                {
                    string[] Values = Sheet_HoldDict[sheetNumber];

                    // Set English note.
                    Transaction t = new Transaction(doc, "Translation by ID");
                    t.Start();
                    // Set english parameter.
                    English_param.Set(Values[0]);
                    Chinese_param.Set(Values[1]);
                    t.Commit();
                    //add center file path to values.
                    // add to update excel from Revit model.
                    string[] array = new string[4];
                    array[0] = sheetNumber;
                    array[1] = Values[0];
                    array[2] = Values[1];
                    array[3] = centralFilePath;
                    HoldIsFoundInRevit.Add(array);
                    SheetsToDeleteFromExcel.Add(sheetNumber, Values);
                }

                //if the dictionary doesn't have translation add to Excel.
                if (!Sheet_IDEnglishChineseDict.ContainsKey(sheetNumber) ||
                    !Sheet_HoldDict.ContainsKey(sheetNumber))
                {
                    if (!SheetNameCompareList.Contains(sheetNumber))
                    {
                        string[] array = new string[4];
                        array[0] = sheetNumber;
                        array[1] = English;
                        array[2] = Chinese;
                        array[3] = centralFilePath;
                        NotTranslated.Add(array);
                    }
                }
            }

            //if dictionary has elements that need to be removed.
            if (SheetsToDeleteFromExcel != null || SheetsToDeleteFromExcel.Count > 0)
                Excel.Delete(path, 3, 1, SheetsToDeleteFromExcel);

            _HoldIsFoundInRevitSheets = HoldIsFoundInRevit;

            return NotTranslated;
        }

        // ***********************************RemoveItemFromExcel***********************************
        public Dictionary<string, string[]> RemoveItemFromExcel(Dictionary<string, string[]> TitleOnSheet_IDEnglishChineseDict, string centralFilePath)
        {
            Dictionary<string, string[]> DeleteFromExcel = new Dictionary<string, string[]>();

            foreach (KeyValuePair<string, string[]> entry in TitleOnSheet_IDEnglishChineseDict)
            {
                String[] values = entry.Value;
                if (centralFilePath == values[2])
                    if (!entry.Key.StartsWith("$")) //Temp views.
                        DeleteFromExcel.Add(entry.Key, entry.Value);
            }
            return DeleteFromExcel;
        }
    }
}