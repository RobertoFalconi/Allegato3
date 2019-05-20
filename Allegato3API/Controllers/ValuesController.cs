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
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        //// POST api/values
        //public HttpResponseMessage Post(HttpRequestMessage request)
        //{
        //    string result;
        //    try
        //    {
        //        result = request.Content.ReadAsStringAsync().Result;
        //        if (result == null)
        //        {
        //            return new HttpResponseMessage(HttpStatusCode.BadRequest);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return new HttpResponseMessage(HttpStatusCode.BadRequest);
        //    }

        //    // POST business logic code goes here

        //    return new HttpResponseMessage(HttpStatusCode.Accepted);
        //}

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
            //files.Nome = result.FirstOrDefault(); TODO: trovare il nome che è la prima chiave
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
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
