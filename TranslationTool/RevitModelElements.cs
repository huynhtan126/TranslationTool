using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TranslationTool
{
    internal class RevitModelElements
    {
        //***********************************GetFamilySymbols***********************************
        public List<FamilyInstance> GetFamilyInstance(Document doc, string name, BuiltInCategory category)
        {
            List<FamilyInstance> List_FamilyInstance = new List<FamilyInstance>();

            ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(FamilyInstance));
            // Category filter
            ElementCategoryFilter Categoryfilter = new ElementCategoryFilter(category);
            // Instance filter
            LogicalAndFilter InstancesFilter = new LogicalAndFilter(familyInstanceFilter, Categoryfilter);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            // Colletion Array of Elements
            ICollection<Element> Elements = collector.WherePasses(InstancesFilter).ToElements();

            foreach (Element e in Elements)
            {
                FamilyInstance familyInstance = e as FamilyInstance;

                if (null != familyInstance)
                {
                    try
                    {
                        if (familyInstance.Name.Contains(name))
                            List_FamilyInstance.Add(familyInstance);
                    }
                    catch (Exception ex)
                    {
                        string x = ex.Message;
                    }
                }
            }
            return List_FamilyInstance;
        }

        //***********************************FindFamilyInstanceByName***********************************
        public FamilyInstance FindFamilyInstanceByName(Document doc, Type targetType, string targetName)
        {
            Element element = new FilteredElementCollector(doc)
              .OfClass(targetType)
              .FirstOrDefault<Element>(
                e => e.Name.Equals(targetName));

            FamilyInstance familyInstance = element as FamilyInstance;

            return familyInstance;
        }

        //***********************************GetAllViewSheets***********************************
        public List<ViewSheet> GetAllSheets(Document doc)
        {
            List<Element> viewElems = new List<Element>();
            List<ViewSheet> Sheets = new List<ViewSheet>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            viewElems.AddRange(collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements());

            foreach (Element e in viewElems)
            {
                if (e is ViewSheet)
                    Sheets.Add(e as ViewSheet);
            }
            return Sheets;
        }

        //***********************************GetLevels***********************************
        public List<Level> GetLevels(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> levels = collector.OfClass(typeof(Level)).ToElements();
            List<Level> List_levels = new List<Level>();
            foreach (Level level in levels)
            {
                List_levels.Add(level);
            }
            return List_levels;
        }

        //***********************************GetAllViews***********************************
        public List<Autodesk.Revit.DB.View> GetAllViews(Document doc)
        {
            FilteredElementCollector viewcollector = new FilteredElementCollector(doc);
            viewcollector.OfClass(typeof(Autodesk.Revit.DB.View));
            ICollection<Element> List_ViewElements = viewcollector.ToElements();
            List<Autodesk.Revit.DB.View> List_Views = new List<Autodesk.Revit.DB.View>();

            foreach (Autodesk.Revit.DB.View v in List_ViewElements)
            {
                List_Views.Add(v);
            }
            return List_Views;
        }

        //***********************************GetParameterValue***********************************
        public string GetParameterValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    //get value with unit, AsDouble() can get value without unit
                    return parameter.AsValueString();

                case StorageType.ElementId:
                    return parameter.AsElementId().IntegerValue.ToString();

                case StorageType.Integer:
                    //get value with unit, AsInteger() can get value without unit
                    return parameter.AsValueString();

                case StorageType.None:
                    return parameter.AsValueString();

                case StorageType.String:
                    return parameter.AsString();

                default:
                    return "";
            }
        }

        //TODO add random string.
        private static Random random = new Random();

        //***********************************RandomString***********************************
        public string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}