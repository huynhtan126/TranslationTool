using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;

namespace TranslationDatabase
{
    class TranslationProgram
    {
        public void Update()
        {
            MongoDBRepository mongoDBRepository = new MongoDBRepository();
            csvCollection csv = new csvCollection();
            EnglishChinese obj1 = new EnglishChinese();
            obj1.Id = "001";
            obj1.English = "EnglishTest1";
            obj1.Chinese = "ChineseTest1";

            mongoDBRepository.Collectoin_EnglishToChinese.InsertOne(obj1);

            //IEnumerable<string[]> csvCollection = csv.GetCsvTranslation();
            //foreach (string[] s in csvCollection)
           
        }
    }
}
