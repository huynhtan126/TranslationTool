using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TranslationWebDatabase
{
    public class EnglishChinese
    {
        [BsonElement("ItemId")]
        public string ItemId { get; set; }
    }
}