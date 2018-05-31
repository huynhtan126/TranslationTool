using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;

namespace TranslationDatabase
{
    class MongoDBRepository
    {
        private IMongoDatabase database;

        public MongoDBRepository()
        {
            MongoClient client = new MongoClient();
            this.database = client.GetDatabase("Database_EnglishToChinese");
            var collection = database.GetCollection<EnglishChinese>("Collection_EnglishToChinese");
        }

        public IMongoCollection<EnglishChinese> Collectoin_EnglishToChinese
        {
            get
            {
                return database.GetCollection<EnglishChinese>("Collectoin_EnglishToChinese");
            }
        }
    }
}
