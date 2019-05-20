using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Allegato3API.Models;

namespace Allegato3API.Controllers
{
    public class FilesController : Controller
    {
        private Allegato3APIContext db = new Allegato3APIContext();

        // GET: Files
        public async Task<ActionResult> Index()
        {
            return View(await db.Files.ToListAsync());
        }

        // GET: Files/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Files files = await db.Files.FindAsync(id);
            if (files == null)
            {
                return HttpNotFound();
            }
            return View(files);
        }

        // GET: Files/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Files/Create
        // Per proteggere da attacchi di overposting, abilitare le proprietà a cui eseguire il binding. 
        // Per ulteriori dettagli, vedere https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "FileID,JsonString")] Files files)
        {
            if (ModelState.IsValid)
            {
                db.Files.Add(files);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(files);
        }

        // GET: Files/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Files files = await db.Files.FindAsync(id);
            if (files == null)
            {
                return HttpNotFound();
            }
            return View(files);
        }

        // POST: Files/Edit/5
        // Per proteggere da attacchi di overposting, abilitare le proprietà a cui eseguire il binding. 
        // Per ulteriori dettagli, vedere https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "FileID,JsonString")] Files files)
        {
            if (ModelState.IsValid)
            {
                db.Entry(files).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(files);
        }

        // GET: Files/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Files files = await db.Files.FindAsync(id);
            if (files == null)
            {
                return HttpNotFound();
            }
            return View(files);
        }

        // POST: Files/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            Files files = await db.Files.FindAsync(id);
            db.Files.Remove(files);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
