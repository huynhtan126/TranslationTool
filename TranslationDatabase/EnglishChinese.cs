using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;

namespace TranslationDatabase
{
    class EnglishChinese
    {
        public string ItemId { get; set; }

        public string English { get; set; }

        public string Chinese { get; set; }
    }
}
