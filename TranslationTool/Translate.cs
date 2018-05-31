using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TranslationTool
{
    internal class Translate
    {
        //Viewport titles field.
        public List<string[]> _UpdateExcelTranslationGenericAnno { get; set; }
        public List<string[]> _NotTranslatedTitleOnSheet { get; set; }
        public List<string[]> _NotTranslatedSheets { get; set; }
        public List<string[]> _HoldIsFoundInRevitSheets { get; set; }
        public List<string[]> _HoldIsFoundInRevitTitleOnSheet { get; set; }

        //***********************************StartTranslation***********************************
        public void StartTranslation(Document doc, Application App, string path, string centralFilePath)
        {
            ExcelSheet Excel = new ExcelSheet();
            //Check Excel file for duplicate items to be deleted.
            Excel.Delete(path, 1, 2);

            //Dictionary of Annotations.  Key as English and value as Chinese.
            Dictionary<string, string> Excel_Anno_DictioinaryEnglishAndChinese = Excel.Read(path, 1);
            List<string> Excel_AnnotationCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            Dictionary<string, string> Excel_Anno_DictionaryIDandEnglishDictionary = Excel._IDandEnglishDictionary;

            //Dicationary of title on sheets. Key as English and value as Chinese.
            Dictionary<string, string> TitleOnSheetsDictionary = Excel.Read(path, 2);
            List<string> TitleOnSheetCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> TitleOnSheet_IDEnglishChineseDict = Excel._IDandArrayDictioinary;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> TitleOnSheet_HoldDict = Excel._HoldDictioinary;

            //Dictionary of Sheet names. Key as English and value as Chinese.
            Dictionary<string, string> SheetsNameDictionary = Excel.Read(path, 3);
            List<string> SheetNameCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> Sheet_IDEnglishChineseDict = Excel._IDandArrayDictioinary;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> Sheet_HoldDict = Excel._HoldDictioinary;

            //Annotations translation.

            //Update ANNOTATION
            TranslateGenericAnnotations TGA = new TranslateGenericAnnotations();
            if (Excel_Anno_DictionaryIDandEnglishDictionary != null)
                TGA.SetEnglishAnnotationByID(doc, App, 
                    Excel_Anno_DictionaryIDandEnglishDictionary);

            List<string[]> UpdateExcelTranslationGenericAnno = TGA.GenericAnnotationTranslation(doc, App,
                Excel_Anno_DictioinaryEnglishAndChinese,
                Excel_Anno_DictionaryIDandEnglishDictionary, 
                Excel_AnnotationCompareList);

            //Update TITLE ON SHEETS
            TranslationTitleOnSheet TTOS = new TranslationTitleOnSheet();
            if (TitleOnSheet_IDEnglishChineseDict != null)
                _NotTranslatedTitleOnSheet = TTOS.SetEnglishTitleOnSheetByID(doc,
                    TitleOnSheet_IDEnglishChineseDict,
                    TitleOnSheet_HoldDict,
                    TitleOnSheetCompareList,
                    Excel, path, centralFilePath);

            //Update SHEET.
            TranslateSheets TS = new TranslateSheets();
            _NotTranslatedSheets = TS.SheetTranslation(doc,
                Sheet_IDEnglishChineseDict,
                Sheet_HoldDict,
                SheetNameCompareList,
                path, centralFilePath);

            //Check Excel file for duplicate  to be deleted.
            //Excel.Delete(path, 2, 1);
            //Check Excel file for duplicate items to be deleted.
            //Excel.Delete(path, 3, 1);

            _HoldIsFoundInRevitSheets = TS._HoldIsFoundInRevitSheets;
            _HoldIsFoundInRevitTitleOnSheet = TTOS._HoldIsFoundInRevitTitleOnSheet;
            _UpdateExcelTranslationGenericAnno = UpdateExcelTranslationGenericAnno;
        }        
    }
}