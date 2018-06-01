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
        static void Main(string[] args)
        {
            

            MongoDBRepository mongoDBRepository = new MongoDBRepository();
            

            csvCollection csv = new csvCollection();
            IEnumerable<string[]> csvCollection = csv.GetCsvTranslation();
            foreach (string[] s in csvCollection)
            {
                EnglishChinese translation = new EnglishChinese
                {
                    ItemId = s[0],
                    English = s[1],
                    Chinese = s[2]
                };
                Console.WriteLine(translation);
                //mongoDBRepository.Collectoin_EnglishToChinese.InsertOne(translation);
            }

        }
    }
}
