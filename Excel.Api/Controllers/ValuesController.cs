using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelNpoi;
using System.IO;

namespace Excel.Api.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {
        // GET api/values
       

        // GET api/values/5
        [HttpGet()]
        public async Task<IActionResult> Get()
        {
            var service = new Service();
            var stream = await service.GenerateXlsx();
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
        }

        [HttpGet("legacy")]
        public async Task<IActionResult> GetLegacy()
        {
            var service = new Service();
            var stream = await service.GenerateXlsx();
            return File(stream, "application/vnd.ms-excel", "xls");
        }

    }
}
