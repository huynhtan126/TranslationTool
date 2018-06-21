using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TranslationWebDatabase.App_Start;
using TranslationWebDatabase.Controllers;
using TranslationWebDatabase.Models;

namespace TranslationTool
{
    class MongoDBCollections
    {
        private MongoContext Collections()
        {
             MongoContext mongoContext = new MongoContext();
             return mongoContext;
        }

        public Dictionary<string, string[]> Read()
        {
            MongoContext mongoContext = Collections();
            List<TranslationE2C_Model> translationE2C_Model = mongoContext.findAll();

            Dictionary<string, string[]> annotationDatabase = new Dictionary<string,string[]>();

            foreach (TranslationE2C_Model model in translationE2C_Model)
            {

                string[] data = new string[3];
                data[0] = model.English;
                data[1] = model.Chinese;
                data[2] = "";
                annotationDatabase.Add(model.Key, data);
            }

            return annotationDatabase;
        }
    }
}
