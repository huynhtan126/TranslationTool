using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace TranslationTool
{
    class RevitParametersCheck
    {
        public void CreateSharedParameters(Document doc, Application app, BuiltInCategory builtInCategory, 
            string groupName ,string ParameterName)
        {
            Category Category = doc.Settings.Categories.get_Item(builtInCategory);
            CategorySet categorySet = app.Create.NewCategorySet();
            categorySet.Insert(Category);

            string originalFile = app.SharedParametersFilename;
            string tempFile = @"X:\Revit\Revit Support\SOM-Structural Parameters.txt";

            try
            {
                app.SharedParametersFilename = tempFile;

                DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();

                foreach (DefinitionGroup dg in sharedParameterFile.Groups)
                {
                    if (dg.Name == groupName)
                    {
                        ExternalDefinition externalDefinition = dg.Definitions.get_Item(ParameterName) as ExternalDefinition;

                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Add Translation Shared Parameters");
                            //parameter binding 
                            InstanceBinding newIB = app.Create.NewInstanceBinding(categorySet);
                            //parameter group to Identity Data in Revit Parameters
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_IDENTITY_DATA);
                            doc.Regenerate();
                            t.Commit();
                        }
                    }
                }
            }
            catch { }
            finally
            {
                //reset to original file
                app.SharedParametersFilename = originalFile;
            }
        }
    }
}
