using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TranslationTool
{
    internal class TranslateGenericAnnotations
    {
        //***********************************GenericAnnotationTranslation***********************************
        public void SetEnglishAnnotationByID(Document doc, Application app,
            Dictionary<string, string> Anno_IDandEnglishDictionary)
        {
            RevitModelElements revit = new RevitModelElements();

            List<FamilyInstance> GenericAnnotations = revit.GetFamilyInstance(doc, "Translation", BuiltInCategory.OST_GenericAnnotation);

            foreach (Element e in GenericAnnotations)
            {
                try
                {
                    //Get the parameters from the project
                    Parameter somID_param = e.LookupParameter("SOM ID");
                    Parameter English_param = e.LookupParameter("ENGLISH");

                    //Check if parameter exist in project.
                    if (somID_param == null)
                        somID_param = CheckIfSharedParamterExist(doc, app, e, "DYNAMO AND ADD-IN", "SOM ID");
                    if (English_param == null)
                        English_param = CheckIfSharedParamterExist(doc, app, e, "TRANSLATION", "ENGLISH");

                    //Get parameter values
                    string id = revit.GetParameterValue(somID_param);
                    string English = revit.GetParameterValue(English_param);

                    if (id != "")
                    {
                        //Check if database dictionary contains the Revit Id.
                        if (Anno_IDandEnglishDictionary.ContainsKey(id))
                        {
                            string EnglishValue = Anno_IDandEnglishDictionary[id];

                            // Set English note.
                            Transaction t = new Transaction(doc, "English Note");
                            t.Start();
                            // Set english parameter.
                            English_param.Set(EnglishValue);
                            t.Commit();
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Set English Annotation by ID Exception raised - " + ex.Message);
                }
            }
        }

        //***********************************GenericAnnotationTranslation***********************************
        public List<string[]> GenericAnnotationTranslation(Document doc, Application app,
            Dictionary<string, string> Excel_Anno_DictioinaryEnglishAndChinese,
            Dictionary<string, string> Excel_Anno_DictionaryIDandEnglishDictionary,
            List<string> AnnotationCompareList)
        {
            RevitModelElements revit = new RevitModelElements();
            List<FamilyInstance> GenericAnnotations = revit.GetFamilyInstance(doc, "Translation", BuiltInCategory.OST_GenericAnnotation);

            List<string[]> UpdateExcelTranslation = new List<string[]>();

            foreach (Element e in GenericAnnotations)
            {
                //Get the parameters from the project
                Parameter somID_param = e.LookupParameter("SOM ID");
                Parameter English_param = e.LookupParameter("ENGLISH");
                Parameter Chinese_param = e.LookupParameter("CHINESE");

                //Check if parameter exist in project.
                if (somID_param == null)
                    somID_param = CheckIfSharedParamterExist(doc, app, e, "DYNAMO AND ADD-IN", "SOM ID");
                if (English_param == null)
                    English_param = CheckIfSharedParamterExist(doc, app, e, "TRANSLATION", "ENGLISH");
                if (Chinese_param == null)
                    Chinese_param = CheckIfSharedParamterExist(doc, app, e, "TRANSLATION", "CHINESE");

                // Revit parameter values.
                string id = revit.GetParameterValue(somID_param);
                string English = revit.GetParameterValue(English_param);
                string Chinese = revit.GetParameterValue(Chinese_param);

                if (Chinese == "" || Chinese == null)
                {
                    if (English != "" || English != null)
                    {
                        if (Excel_Anno_DictioinaryEnglishAndChinese.ContainsKey(English))
                        {
                            string ChineseValue = Excel_Anno_DictioinaryEnglishAndChinese[English];

                            // Set Chinese translation.
                            Transaction t = new Transaction(doc, "Translation");
                            t.Start();
                            // set in transaction.
                            Chinese_param.Set(ChineseValue);
                            // set id if set to none.
                            if (id == "" || id == null)
                            {
                                //Check if key value exist in dictionary.
                                string key = Excel_Anno_DictionaryIDandEnglishDictionary.FirstOrDefault(x => x.Value == English).Key;
                                if (key != "" || key != null)
                                    somID_param.Set(key);
                                //Generate a key if doesn't have one.
                                if (key == "" || key == null)
                                {
                                    //Add new ID if no Id is assigned.
                                    id = RandomKeyId(Excel_Anno_DictionaryIDandEnglishDictionary, revit);
                                    somID_param.Set(id);
                                    string[] array = new string[4];
                                    array[0] = id;
                                    array[1] = English;
                                    array[2] = ChineseValue;
                                    array[3] = "";
                                    UpdateExcelTranslation.Add(array);
                                }
                            }
                            t.Commit();
                        }
                    }

                    // Check if the english has value.
                    if (English != "" || English != null)
                    {
                        //if the dictionary doesn't have translation add to Excel.
                        if (!Excel_Anno_DictioinaryEnglishAndChinese.ContainsKey(English))
                        {
                            if (!AnnotationCompareList.Contains(English))
                            {
                                if (id == "" || id == null)
                                {
                                    //Add new ID if no Id is assigned.
                                    id = RandomKeyId(Excel_Anno_DictionaryIDandEnglishDictionary, revit);
                                }

                                string[] array = new string[4];
                                array[0] = id;
                                array[1] = English;
                                array[2] = "";
                                array[3] = "";

                                UpdateExcelTranslation.Add(array);
                            }
                        }
                    }
                }
            }
            return UpdateExcelTranslation;
        }

        //***********************************CheckIfSharedParamterExist***********************************
        public Parameter CheckIfSharedParamterExist(Document doc, Application app, Element e,
            string groupName, string parameterName)
        {
            RevitParametersCheck revitParameters = new RevitParametersCheck();
            Parameter newParameter;

            revitParameters.CreateSharedParameters(doc, app, BuiltInCategory.OST_GenericAnnotation,
                groupName, parameterName);
            newParameter = e.LookupParameter(parameterName);
            return newParameter;
        }

        //***********************************NewKeyId***********************************
        public string RandomKeyId(Dictionary<string, string> Anno_DictionaryIDandEnglishDictionary,
            RevitModelElements revit)
        {
            // Add new ID
            string newID = "1";
            // if the newID if found in the dictionary create a new Id.
            while (Anno_DictionaryIDandEnglishDictionary.ContainsKey(newID))
            {
                newID = revit.RandomString(5);
            }
            return newID;
        }
    }
}