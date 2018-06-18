using TranslationWebDatabase.App_Start;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TranslationWebDatabase.Models;
using MongoDB.Bson;
using MongoDB.Driver.Builders;
using MongoDB.Driver;

namespace TranslationWebDatabase.Controllers
{
    public class TranslationController : Controller
    {
        MongoContext _dbContext;
        public TranslationController()
        {
            _dbContext = new MongoContext();
        }

        //
        // GET: /Translation/
        public ActionResult Index()
        {
            var translationDetails = new List<TranslationE2C_Model>();

            translationDetails = _dbContext.findAll().ToList();

            return View(translationDetails);
        }

        //
        // GET: /Translation/Details/5
        public ActionResult Details(string id)
        {
            var transationId = new TranslationE2C_Model();
            transationId = _dbContext.find(id);
            return View(transationId);
        }

        //
        // GET: /Translation/Create
        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Translation/Create
        [HttpPost]
        public ActionResult Create(TranslationE2C_Model translationE2C_model)
        {
            IMongoCollection<TranslationE2C_Model> document = _dbContext._collection;
            
            //var query = Query.And(Query.EQ("ItemId", translationE2C_model.ItemId), Query.EQ("English", translationE2C_model.English));

            //var count = document.FindAs<BsonDocument>(query).Count();

            //if (count == 0)
            //{

            //}
            //else
            //{
            //    TempData["Message"] = "Translation Already Exist";
            //    return View("Create", translationE2C_model);
            //}

            document.InsertOne(translationE2C_model);

            return RedirectToAction("Index");
        }

        //
        // GET: /Translation/Edit/5
        public ActionResult Edit(string id)
        {
            var translateId = new TranslationE2C_Model();
            translateId = _dbContext.find(id);

            if (translateId != null)
            {
                return View(translateId);
            }
            return RedirectToAction("Index");
        } 

        //
        // POST: /Translation/Edit/5
        [HttpPost]
        public ActionResult Edit(string id, TranslationE2C_Model translationModel)
        {
            try
            {
                _dbContext.update(translationModel);

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            } 
        }

        //
        // GET: /Translation/Delete/5
        public ActionResult Delete(string id)
        {
            var translateId = new TranslationE2C_Model();
            translateId = _dbContext.find(id);

            if(translateId != null)
            {
                return View(translateId);
            }
            return RedirectToAction("Index");
        }

        //
        // POST: /Translation/Delete/5
        [HttpPost]
        public ActionResult Delete(string id, TranslationE2C_Model translationModel)
        {
            try
            {
                _dbContext.delete(id);

                return RedirectToAction("Index");
            }
            catch
            {
                return RedirectToAction("Index");
            }
        }
    }
}
