using MongoDB.Driver;
using System;
using System.Configuration;

namespace TranslationWebDatabase.App_Start
{
    public class MongoContext
    {

        MongoClient _client;
        public IMongoDatabase _database;

        public MongoContext()
        {
            //Reading credentials from Web.config file
            var MongoDatabaseName = ConfigurationManager.AppSettings["MongoDatabaseName"];
            var MongoUsername = ConfigurationManager.AppSettings["MongoUsername"];
            var MongoPassword = ConfigurationManager.AppSettings["MongoPassword"];
            var MongoPort = ConfigurationManager.AppSettings["MongoPort"];  //27017  
            var MongoHost = ConfigurationManager.AppSettings["MongoHost"];

            //Creating credentials 
            var credential = MongoCredential.CreateMongoCRCredential(
                MongoDatabaseName, MongoUsername, MongoPassword);

            //Creating MongoClientSettings
            var settings = new MongoClientSettings
            {
                Credentials = new { credential },
                Server = new MongoServerAddress(MongoHost, Convert.ToInt32(MongoPort))
            };

            _client = new MongoClient(settings);
            //_server = _client.GetServer();
            _database = _client.GetDatabase(MongoDatabaseName);
        }
    }
}