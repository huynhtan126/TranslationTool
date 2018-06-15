using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TranslationTool
{
    internal class TranslationTitleOnSheet
    {
        public List<string[]> _HoldIsFoundInRevitTitleOnSheet { get; set; }

        //***********************************SetEnglishTitleOnSheetByID***********************************
        public List<string[]> SetEnglishTitleOnSheetByID(Document doc,
            Dictionary<string, string[]> TitleOnSheet_IDEnglishChineseDict,
            Dictionary<string, string[]> TitleOnSheet_HoldDict,
            List<string> TitleOnSheetCompareList,
            ExcelSheet Excel, string path, string centralFilePath)
        {
            RevitModelElements revit = new RevitModelElements();
            List<ViewSheet> sheets = revit.GetAllSheets(doc);
            //Filter to check only the views in current docuement.
            Dictionary<string, string[]> TitleOnSheetsToDeleteFromExcel = RemoveItemFromExcel(TitleOnSheet_IDEnglishChineseDict, centralFilePath);

            List<string[]> NotTranslated = new List<string[]>();
            List<string[]> HoldIsFoundInRevit = new List<string[]>();

            foreach (ViewSheet vs in sheets)
            {
                try
                {
                    ICollection<ElementId> listViews = vs.GetAllViewports();
                    foreach (ElementId Elmids in listViews)
                    {
                        Element elementViewports = doc.GetElement(Elmids);
                        Viewport viewport = elementViewports as Viewport;
                        ElementId viewId = viewport.ViewId;
                        Element elementView = doc.GetElement(viewId);
                        Autodesk.Revit.DB.View v = elementView as Autodesk.Revit.DB.View;

                        //Combine the sheet number and detail number to make a id
                        string sheetNumber = viewport.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER).AsString();
                        string detailNumber = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString();
                        string viewportID = sheetNumber + "-" + detailNumber;

                        var viewType = v.ViewType.ToString();
                        Parameter titleOnSheetParam = v.LookupParameter("Title on Sheet");
                        Parameter titleOnSheetChineseParam = v.LookupParameter("Title on Sheet (Chinese)");
                        string RvtChinese = revit.GetParameterValue(titleOnSheetChineseParam);
                        string RvtEnglish = revit.GetParameterValue(titleOnSheetParam);

                        //Check if the viewport is on a Z sheet.
                        //If viewport is on Z sheet skip over it.
                        if (viewportID.StartsWith("z") || viewportID.StartsWith("Z"))
                            continue;

                        //Check if the title is in the database
                        if (TitleOnSheet_IDEnglishChineseDict.ContainsKey(viewportID))
                        {
                            string[] Values = TitleOnSheet_IDEnglishChineseDict[viewportID];
                            // If the Revit Chinese parameter is empty
                            // set the value from the database
                            if (RvtChinese == "" || RvtChinese == null)
                            {
                                // Set English note.
                                Transaction t = new Transaction(doc, "Translation by ID");
                                t.Start();
                                // Set english parameter.
                                titleOnSheetParam.Set(Values[0]);
                                titleOnSheetChineseParam.Set(Values[1]);
                                t.Commit();
                            }
                            //If view exists in Revit remove from list.
                            TitleOnSheetsToDeleteFromExcel.Remove(viewportID);
                        }

                        //Check if Project missing, but translation exists
                        //in Revit needs to be translated and project added to database.
                        if (TitleOnSheet_HoldDict.ContainsKey(viewportID))
                        {
                            string[] Values = TitleOnSheet_HoldDict[viewportID];

                            // Set English note.
                            Transaction t = new Transaction(doc, "Translation by ID");
                            t.Start();
                            // Set english parameter.
                            titleOnSheetParam.Set(Values[0]);
                            titleOnSheetChineseParam.Set(Values[1]);
                            t.Commit();
                            // add center file path to values.
                            // add to update excel from Revit model.
                            string[] array = new string[4];
                            array[0] = viewportID;
                            array[1] = Values[0];
                            array[2] = Values[1];
                            array[3] = centralFilePath;
                            HoldIsFoundInRevit.Add(array);
                            TitleOnSheetsToDeleteFromExcel.Add(viewportID, Values);
                        }

                        //if the dictionary doesn't have translation add to Excel.
                        if (TitleOnSheet_IDEnglishChineseDict.ContainsKey(viewportID) == false)
                            if (TitleOnSheet_HoldDict.ContainsKey(viewportID) == false)
                                if (!TitleOnSheetCompareList.Contains(viewportID))
                                {
                                    string[] array = new string[4];
                                    array[0] = viewportID;
                                    array[1] = RvtEnglish;
                                    array[2] = RvtChinese;
                                    array[3] = centralFilePath;
                                    NotTranslated.Add(array);
                                }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Set English Title On Sheet by ID Exception raised - " + ex.Message);
                }
            }

            //if dictionary has elements that need to be removed.
            if (TitleOnSheetsToDeleteFromExcel != null || TitleOnSheetsToDeleteFromExcel.Count > 0)
                Excel.Delete(path, 2, 1, TitleOnSheetsToDeleteFromExcel);

            _HoldIsFoundInRevitTitleOnSheet = HoldIsFoundInRevit;
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
                    DeleteFromExcel.Add(entry.Key, entry.Value);
            }
            return DeleteFromExcel;
        }
    }
}