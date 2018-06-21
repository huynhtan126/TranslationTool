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

        public List<string[]> UpdateRevit_AnnotationsTranslation(Document doc, Application app,
            Dictionary<string, string[]> AnnotationDatabase)
        {
            //Dictionary to update the database
            List<string[]> UpdateAnnotationDatabase = new List<string[]>();

            RevitModelElements revitElements = new RevitModelElements();
            List<FamilyInstance> annotationFamilyInstance = revitElements.GetFamilyInstance(doc, "Translation", BuiltInCategory.OST_GenericAnnotation);

            foreach (Element e in annotationFamilyInstance)
            {
                //Get the parameters from the project
                Parameter somID_param = e.LookupParameter("SOM ID");
                Parameter English_param = e.LookupParameter("ENGLISH");
                Parameter Chinese_param = e.LookupParameter("CHINESE");

                // Revit parameter values.
                string SOM_id = revitElements.GetParameterValue(somID_param);
                string English = revitElements.GetParameterValue(English_param);
                string Chinese = revitElements.GetParameterValue(Chinese_param);

                //UPDATE USING KEY
                if (SOM_id != "")
                    if (SOM_id != null)
                    {
                        {
                            if (AnnotationDatabase.ContainsKey(SOM_id)) // Find key in dictionary
                            {
                                // Get values in database dictionary. 
                                string[] Values = AnnotationDatabase[SOM_id];
                                // Set Revit parameters 
                                Transaction t = new Transaction(doc, "Translation");
                                t.Start();
                                // If English is empty set from database
                                English_param.Set(Values[0]);
                                // If Chinese is empty set from database
                                Chinese_param.Set(Values[1]);

                                t.Commit();
                                continue;
                            }
                        }
                    }

                //UPDATE WITH ENGLISH ONLY
                string key = AnnotationDatabase.FirstOrDefault(x => x.Value[0] == English).Key;

                if (key != "")
                    if (key != null)
                    {
                        {
                            if (AnnotationDatabase.ContainsKey(key)) // Find key in dictionary
                            {
                                // Get values in database dictionary. 
                                string[] Values = AnnotationDatabase[key];
                                // Set Revit parameters 
                                Transaction t = new Transaction(doc, "Translation");
                                t.Start();
                                // If Key is empty set from database 
                                if (SOM_id == "" || SOM_id == null)
                                    somID_param.Set(key.ToString());
                                // If English is empty set from database
                                if (English == "" || English == null)
                                    English_param.Set(Values[0]);
                                // If Chinese is empty set from database
                                if (Chinese == "" || Chinese == null)
                                    Chinese_param.Set(Values[1]);

                                t.Commit();
                                continue;
                            }
                        }
                    }

                //CREATE RANDOM KEY AND UPDATE TO DATABASE
                if (key == "" || key == null)
                {
                    {
                        // Generate Key and set in Revit.
                        SOM_id = RandomKeyId(AnnotationDatabase, revitElements);

                        // Set Id in Revit. 
                        Transaction t = new Transaction(doc, "Translation");
                        t.Start();
                        somID_param.Set(SOM_id);
                        t.Commit();

                        // Add to update excel file. 
                        if (English != "" || English != null)
                        {
                            string[] array = new string[4];
                            array[0] = SOM_id;
                            array[1] = English;
                            array[2] = Chinese;
                            array[3] = "";
                            UpdateAnnotationDatabase.Add(array);
                        }
                        continue;
                    }
                }

            }
            return UpdateAnnotationDatabase;
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

        //***********************************NewKeyId***********************************
        public string RandomKeyId(Dictionary<string, string[]> Anno_DictionaryIDandEnglishDictionary,
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