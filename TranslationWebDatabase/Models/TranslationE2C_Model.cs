using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace TranslationWebDatabase.Models
{
    public class TranslationE2C_Model
    {
        [BsonId]
        public ObjectId Id { get; set; }

        [BsonElement("ItemId")]
        public string ItemId { get; set; }

        [BsonElement("English")]
        public string English { get; set; }

        [BsonElement("Chinese")]
        public string Chinese { get; set; }
    }
}