//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Data.Entity;
//using System.Data.Entity.Infrastructure;
//using System.Linq;
//using System.Net;
//using System.Net.Http;
//using System.Threading.Tasks;
//using System.Web.Http;
//using System.Web.Http.ModelBinding;
//using System.Web.Http.OData;
//using System.Web.Http.OData.Routing;
//using Allegato3API.Models;

//namespace Allegato3API.Controllers
//{
//    /*
//    Per aggiungere una route relativa a questo controller, può essere necessario apportare altre modifiche alla classe WebApiConfig. Unire queste istruzioni nel metodo Register della classe WebApiConfig. Tenere presente che per gli URL OData viene fatta distinzione tra maiuscole e minuscole.

//    using System.Web.Http.OData.Builder;
//    using System.Web.Http.OData.Extensions;
//    using Allegato3API.Models;
//    ODataConventionModelBuilder builder = new ODataConventionModelBuilder();
//    builder.EntitySet<Files>("FilesApi");
//    config.Routes.MapODataServiceRoute("odata", "odata", builder.GetEdmModel());
//    */
//    public class FilesApiController : ODataController
//    {
//        private Allegato3APIContext db = new Allegato3APIContext();

//        // GET: odata/FilesApi
//        [EnableQuery]
//        public IQueryable<Files> GetFilesApi()
//        {
//            return db.Files;
//        }

//        //// GET: odata/FilesApi(5)
//        //[EnableQuery]
//        //public SingleResult<Files> GetFiles([FromODataUri] int key)
//        //{
//        //    //return SingleResult.Create(db.Files.Where(files => files.FileID == key));
//        //}

//        // PUT: odata/FilesApi(5)
//        public async Task<IHttpActionResult> Put([FromODataUri] int key, Delta<Files> patch)
//        {
//            Validate(patch.GetEntity());

//            if (!ModelState.IsValid)
//            {
//                return BadRequest(ModelState);
//            }

//            Files files = await db.Files.FindAsync(key);
//            if (files == null)
//            {
//                return NotFound();
//            }

//            patch.Put(files);

//            try
//            {
//                await db.SaveChangesAsync();
//            }
//            catch (DbUpdateConcurrencyException)
//            {
//                if (!FilesExists(key))
//                {
//                    return NotFound();
//                }
//                else
//                {
//                    throw;
//                }
//            }

//            return Updated(files);
//        }

//        // POST: odata/FilesApi
//        public async Task<IHttpActionResult> Post(Files files)
//        {
//            if (!ModelState.IsValid)
//            {
//                return BadRequest(ModelState);
//            }

//            db.Files.Add(files);
//            await db.SaveChangesAsync();

//            return Created(files);
//        }

//        // PATCH: odata/FilesApi(5)
//        [AcceptVerbs("PATCH", "MERGE")]
//        public async Task<IHttpActionResult> Patch([FromODataUri] int key, Delta<Files> patch)
//        {
//            Validate(patch.GetEntity());

//            if (!ModelState.IsValid)
//            {
//                return BadRequest(ModelState);
//            }

//            Files files = await db.Files.FindAsync(key);
//            if (files == null)
//            {
//                return NotFound();
//            }

//            patch.Patch(files);

//            try
//            {
//                await db.SaveChangesAsync();
//            }
//            catch (DbUpdateConcurrencyException)
//            {
//                if (!FilesExists(key))
//                {
//                    return NotFound();
//                }
//                else
//                {
//                    throw;
//                }
//            }

//            return Updated(files);
//        }

//        // DELETE: odata/FilesApi(5)
//        public async Task<IHttpActionResult> Delete([FromODataUri] int key)
//        {
//            Files files = await db.Files.FindAsync(key);
//            if (files == null)
//            {
//                return NotFound();
//            }

//            db.Files.Remove(files);
//            await db.SaveChangesAsync();

//            return StatusCode(HttpStatusCode.NoContent);
//        }

//        protected override void Dispose(bool disposing)
//        {
//            if (disposing)
//            {
//                db.Dispose();
//            }
//            base.Dispose(disposing);
//        }

//        //private bool FilesExists(int key)
//        //{
//        //    return db.Files.Count(e => e.FileID == key) > 0;
//        //}
//    }
//}
