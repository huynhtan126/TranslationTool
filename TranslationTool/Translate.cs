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
        public Dictionary<string, string[]> _TitleOnSheet_IDEnglishChineseDict { get; set; }
        public Dictionary<string, string[]> _Sheet_IDEnglishChineseDict { get; set; }
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
            #region ANNOTATIONS 
            //Dictionary of Annotations.  Key as English and value as Chinese.
            Dictionary<string, string> Excel_Anno_DictioinaryEnglishAndChinese = Excel.Read(path, 1);
            List<string> Excel_AnnotationCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            Dictionary<string, string> _Anno_IDandEnglishDictionary = Excel._IDandEnglishDictionary;
            #endregion

            #region TITLE ON SHEETS
            //Dicationary of title on sheets. Key as English and value as Chinese.
            Dictionary<string, string> TitleOnSheetsDictionary = Excel.Read(path, 2);
            List<string> TitleOnSheetCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            _TitleOnSheet_IDEnglishChineseDict = Excel._IDandArrayDictioinary;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> TitleOnSheet_HoldDict = Excel._HoldDictioinary;
            #endregion

            #region SHEET NUMBER
            //Dictionary of Sheet names. Key as English and value as Chinese.
            Dictionary<string, string> SheetsNameDictionary = Excel.Read(path, 3);
            List<string> SheetNameCompareList = Excel.CompareList;

            //Dictionary of English words and Ids.
            _Sheet_IDEnglishChineseDict = Excel._IDandArrayDictioinary;

            //Dictionary of English words and Ids.
            Dictionary<string, string[]> Sheet_HoldDict = Excel._HoldDictioinary;
            #endregion


            //Update ANNOTATION
            TranslateGenericAnnotations TGA = new TranslateGenericAnnotations();
            if (_Anno_IDandEnglishDictionary != null)
                TGA.SetEnglishAnnotationByID(doc, App, 
                    _Anno_IDandEnglishDictionary);

            List<string[]> UpdateExcelTranslationGenericAnno = TGA.GenericAnnotationTranslation(doc, App,
                Excel_Anno_DictioinaryEnglishAndChinese,
                _Anno_IDandEnglishDictionary, 
                Excel_AnnotationCompareList);

            //Update TITLE ON SHEETS
            TranslationTitleOnSheet TTOS = new TranslationTitleOnSheet();
            if (_TitleOnSheet_IDEnglishChineseDict != null)
                _NotTranslatedTitleOnSheet = TTOS.SetEnglishTitleOnSheetByID(doc,
                    _TitleOnSheet_IDEnglishChineseDict,
                    TitleOnSheet_HoldDict,
                    TitleOnSheetCompareList,
                    Excel, path, centralFilePath);

            //Update SHEET.
            TranslateSheets TS = new TranslateSheets();
            _NotTranslatedSheets = TS.SheetTranslation(doc,
                _Sheet_IDEnglishChineseDict,
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