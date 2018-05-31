using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslationDatabase
{
    class csvCollection
    {
        public IEnumerable<string[]> GetCsvTranslation()
        {
            string fileName = @"c\SOM-Chinese-Translation.csv";
            var translations = File.ReadLines(fileName).Take(10000);
            return translations.Skip(1).Select(s => s.Split(','));
        }
    }
}
