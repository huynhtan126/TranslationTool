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
        public MongoContext MongoDBCollections()
        {
             MongoContext mongoContext = new MongoContext();
             return mongoContext;
        }

        public List<TranslationE2C_Model> Read()
        {
            MongoContext mongoContext = MongoDBCollections();
            List<TranslationE2C_Model> translationE2C_Model = mongoContext.findAll();

            return translationE2C_Model;
        }
    }
}
