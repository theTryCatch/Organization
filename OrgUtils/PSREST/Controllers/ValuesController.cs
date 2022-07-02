using OrgUtils.PowerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Newtonsoft.Json;

namespace PSREST.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            var a = new PowerShell(new List<string>() { "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1", "mslaptop", "dc1" }, "get-service | select Name, displayName, status", CodeType.Script, null, null);
            foreach (var item in a.BeginInvoke())
            {
                yield return JsonConvert.SerializeObject(item);
            }
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
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
