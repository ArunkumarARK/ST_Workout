using System;
using System.Globalization;
using System.Security.AccessControl;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Newtonsoft.Json;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using SmartTrack.Models;
using System.Text.RegularExpressions;
using SmartTrack.Helper;
using System.Net.Mail;

namespace SmartTrack.Controllers
{
    public class InvoiceController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: Invoice
        public ActionResult Index()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult Billing()
        {
            //Art TAT
            try
            {
                string strquery = "select * from (select pm1.process as 'Core FrontList', pm.process, Concat(p.rate,'_',p.AutoSeqID) as rate from BK_processInfo p, JBM_processmaster pm, JBM_processmaster pm1 where p.processid = pm1.processid and p.SubProcess = pm.processid) as tab1 pivot (max(tab1.rate) for tab1.process in (\"500_core_onshore\", \"500_core_offshore\", \"750_core_onshore\", \"750_core_offshore\", \"1000_core_onshore\", \"1000_core_offshore\", \"1250_core_onshore\", \"1250_core_offshore\", \"1500_core_onshore\", \"1500_core_offshore\", \"1750_core_onshore\", \"1750_core_offshore\", \"2000_core_onshore\", \"2000_core_offshore\")) as tab2";
                DataTable dtTAT = new DataTable();
                dtTAT = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());

                DataSet ds = new DataSet();
                ds.Tables.Add(dtTAT);
                return View(ds);
            }
            catch (Exception)
            {
                return null;
            }
        }
        [SessionExpire]
        public ActionResult UpdateBilling(string BillingData)
        {
            try
            {
                if (BillingData.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";
                    List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(BillingData);
                    if (chkIds.Count > 0)
                    {
                        for (int i = 0; i < chkIds.Count; i++)
                        {
                            string strId = chkIds[i].Split('|')[0];
                            string strval = chkIds[i].Split('|')[1];
                            string strQuery = "";

                            strQuery = "Update BK_ProcessInfo set Rate=" + strval + " where AutoSeqID='" + strId + "'";

                            string strResult = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                        }
                    }
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
    }
}