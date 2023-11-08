using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SmartTrack.Controllers
{
    public class ErrorHandlerController : Controller
    {
        // GET: ErrorHandler
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult NotFound()
        {
            return View();
        }
        public ActionResult SqlExceptionView()
        {
            return View();
        }
        [HandleError(ExceptionType = typeof(SqlException), View = "SqlExceptionView")]
        public string GetClientInfo(string username)
        {
            return "false";
        }
    }
}