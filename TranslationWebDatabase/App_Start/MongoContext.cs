using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using System;
using System.Collections.Generic;
using System.Configuration;
using TranslationWebDatabase.Models;
using System.Linq;

namespace TranslationWebDatabase.App_Start
{
    public class MongoContext
    {
        private MongoClient _client;
        public IMongoDatabase _database;
        public IMongoCollection<TranslationE2C_Model> _collection;

        //***********************************MongoContext***********************************
        public MongoContext()
        {
            //Reading credentials from Web.config file
            var MongoDatabaseName = ConfigurationManager.AppSettings["MongoDatabaseName"];
            var MongoDatabaseHost = ConfigurationManager.AppSettings["MongoDatabaseHost"];

            //Creating credentials 
            _client = new MongoClient(MongoDatabaseHost);
            var db = _client.GetDatabase(MongoDatabaseName);
            _collection = db.GetCollection<TranslationE2C_Model>("translateE2C_model");
        }

        //***********************************findAll***********************************
        public List<TranslationE2C_Model> findAll()
        {
            return _collection.AsQueryable<TranslationE2C_Model>().ToList();
        }

        //***********************************find***********************************
        public TranslationE2C_Model find(string id)
        {
            var translationId = new ObjectId(id);

            return _collection.AsQueryable<TranslationE2C_Model>().SingleOrDefault(x => x.Id == translationId);
        }

        //***********************************create***********************************
        public void create(TranslationE2C_Model translate)
        {
            _collection.InsertOne(translate);
        }

        //***********************************update***********************************
        public void update(TranslationE2C_Model translate)
        {
            _collection.UpdateOne(
                Builders<TranslationE2C_Model>.Filter.Eq("Id", translate.Id),
                Builders<TranslationE2C_Model>.Update
                .Set("ItemId", translate.ItemId)
                .Set("English", translate.English)
                .Set("Chinese", translate.Chinese)
                );
        }

        //***********************************delete***********************************
        public void delete(string id)
        {
            var translationId = new ObjectId(id);
            TranslationE2C_Model translate = _collection.AsQueryable<TranslationE2C_Model>().SingleOrDefault(x => x.Id == translationId);
            _collection.DeleteOne(Builders<TranslationE2C_Model>.Filter.Eq("Id", translate.Id));
        }
    }
}