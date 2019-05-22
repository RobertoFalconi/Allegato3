using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.ModelBinding;
using System.Web.Http.OData;
using System.Web.Http.OData.Routing;
using Allegato3API.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Allegato3API.Controllers
{
    public class ValuesController : ApiController
    {
        private Allegato3APIContext db = new Allegato3APIContext();

        // GET api/values
        public async Task<IEnumerable<string>> Get(HttpRequestMessage request)
        {
            var ret = db.Files.Select(x => x.Nome);
            return ret;
        }

        // GET api/values/5
        public string Get(int id)
        {
            // TODO: implementare la GET
            return "{\"the best prova#Foglio1\":[{\"0\":[{\"Item1\":\"sciau belu\",\"Item2\":16777215.0,\"Item3\":false}]}]}";
        }

        // POST api/values
        public async Task<HttpResponseMessage> Post(HttpRequestMessage request)
        {
            string result;
            try
            {
                result = request.Content.ReadAsStringAsync().Result;
                if (result == null)
                {
                    return new HttpResponseMessage(HttpStatusCode.BadRequest);
                }
            }
            catch (Exception ex)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }

            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(result);
            Files files = new Files();
            files.Nome = jsonObject.Properties().Select(p => p.Name).FirstOrDefault().Split('#')[0];
            files.JsonString = result;

            if (!ModelState.IsValid)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }

            db.Files.Add(files);
            await db.SaveChangesAsync();

            return new HttpResponseMessage(HttpStatusCode.Accepted);
        }

        // PUT api/values/5
        public async Task<HttpResponseMessage> Put(HttpRequestMessage request)
        {
            string result;
            try
            {
                result = request.Content.ReadAsStringAsync().Result;
                if (result == null)
                {
                    return new HttpResponseMessage(HttpStatusCode.BadRequest);
                }
            }
            catch (Exception ex)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }

            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(result);
            Files files = await db.Files.FindAsync(jsonObject.Properties().Select(p => p.Name).FirstOrDefault().Split('#')[0]);

            if (!ModelState.IsValid)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }

            files.JsonString = result;
            db.Entry(files).State = EntityState.Modified;
            await db.SaveChangesAsync();

            return new HttpResponseMessage(HttpStatusCode.Accepted);
        }

        // DELETE api/values/5
        public async Task<HttpResponseMessage> Delete(HttpRequestMessage request)
        {
            string result;
            try
            {
                result = request.Content.ReadAsStringAsync().Result;
                if (result == null)
                {
                    return new HttpResponseMessage(HttpStatusCode.BadRequest);
                }
            }
            catch (Exception ex)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }

            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(result);
            string id = jsonObject.Properties().Select(p => p.Name).FirstOrDefault().Split('#')[0];
            Files files = await db.Files.FindAsync(id);
            db.Files.Remove(files);
            await db.SaveChangesAsync();

            return new HttpResponseMessage(HttpStatusCode.Accepted);

        }
    }
}
