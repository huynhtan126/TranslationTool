using TranslationWebDatabase.App_Start;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TranslationWebDatabase.Models;
using MongoDB.Bson;
using MongoDB.Driver;

namespace TranslationWebDatabase.Controllers
{
    public class TranslationController : Controller : I
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
            return View();
        }

        //
        // GET: /Translation/Details/5
        public ActionResult Details(int id)
        {
            return View();
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
            var document = _dbContext._database.GetCollection<BsonDocument>("translateE2C_model");
            
            var query = Query.And(Query.EQ("ItemId", translationE2C_model.ItemId), Query.EQ("English", translationE2C_model.English));

            var count = document.FindAs<translationE2C_model>(query).Count();

            if (count == 0)
            {
                var result = document.InsertOne(translationE2C_model);
            }
            else
            {
                TempData["Message"] = "Translation Already Exist";
                return View("Create", translationE2C_model);
            }
        }

        //
        // GET: /Translation/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Translation/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Translation/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Translation/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
