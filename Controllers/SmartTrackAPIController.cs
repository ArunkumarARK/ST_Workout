using System;
using System.IO;
using System.Data;
using System.Xml;
using System.Xml.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Net.Mail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Text;

namespace SmartTrack.Controllers
{
    public class SmartTrackAPIController : ApiController
    {

        DataProc DBProc = new DataProc();

        [AllowAnonymous]
        [HttpPost]
        [Route("api/GetJournalList")] 
        public IHttpActionResult GetJournalInfo()
        {
            
            

            return Ok("Success");
        }



            //// GET api/<controller>
            //public IEnumerable<string> Get()
            //{
            //    return new string[] { "value1", "value2" };
            //}

            //// GET api/<controller>/5
            //public string Get(int id)
            //{
            //    return "value";
            //}

            //// POST api/<controller>
            //public void Post([FromBody]string value)
            //{
            //}

            //// PUT api/<controller>/5
            //public void Put(int id, [FromBody]string value)
            //{
            //}

            //// DELETE api/<controller>/5
            //public void Delete(int id)
            //{
            //}
        }
}