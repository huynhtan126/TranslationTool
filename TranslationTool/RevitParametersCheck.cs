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
        public void CreateSharedParameters(Document doc, Application app, BuiltInCategory builtInCategory, string ParameterName)
        {
            Category sheetCategory = doc.Settings.Categories.get_Item(builtInCategory);
            CategorySet categorySet = app.Create.NewCategorySet();
            categorySet.Insert(category);

            string originalFile = app.SharedParametersFilename;
            string tempFile = @"X:\Revit\Revit Support\SOM-Structural Parameters.txt";

            try
            {
                app.SharedParametersFilename = tempFile;

                DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();

                foreach (DefinitionGroup dg in sharedParameterFile.Groups)
                {
                    if (dg.Name == "DYNAMO AND ADD-IN")
                    {
                        ExternalDefinition externalDefinition = dg.Definitions.get_Item("GROUP 1") as ExternalDefinition;

                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Add Shared Parameters");
                            //parameter binding 
                            InstanceBinding newIB = app.Create.NewInstanceBinding(categorySet);
                            //parameter group to text
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA);
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
