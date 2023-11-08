using System;
using System.Globalization;
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
using ReferenceLibrary.IniDet;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Data.OleDb;

namespace SmartTrack.Controllers
{
    public class NasExplorerController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: NasExplorer
        public ActionResult Index()
        {
            return View();
        }
       public ActionResult btnNASExp_Click(string strCommand, string strEmpLogin, string strCeninwExp,string strCenAppExp, string strCenproExp,string strWorkDirExp,string strCeninwDirExpSupport,string strNASExplorer,string strUniqueId)
        {
            //data-Command  -- Folder access command
            //data-CeninwExp -- Ceninw path
            //data-CenproExp -- Cenpro Path
            //data-CenproExp -- CenApp Path
            //data-WorkDirExp -- Working Directory Path
            //data-emp  -- Emp Login ID
            //data-CeninwExpSupport -- Support path
            int intW  = ((Request.Browser.ScreenPixelsWidth) * 2 + 100);
            int intH = ((Request.Browser.ScreenPixelsHeight) * 2 + 100);

            //strCommand = "CNINW|CENPRO|WORKDIR|ROOTACCESS|SUPPORT-RIGHTS|CINF|CIRE|CIUP|CICP|CIPS|CIDL|CNF|CRE|CUP|CCP|CPS|CDL|WNF|WRE|WUP|WCP|WPS|WDL";
            //strEmpLogin = "40385";
            //strCeninwExp = "[$][$]chenasprod[$]kglinw[$]Journals[$]KGLVT[$]INFORMS[$]ORSC[$]Vol00000[$]220093[$]";
            //strCenAppExp = "[$][$]chenasprod[$]kglpropdf[$]ApplicationFiles[$]Journals[$]KGLVT[$]INFORMS[$]ORSC[$]Vol00000[$]220093[$]";
            //strCenproExp = "[$][$]chenasprod[$]kglpro[$]ApplicationFiles[$]Journals[$]KGLVT[$]INFORMS[$]OPRE[$]Vol00000[$]210233[$]";
            //strWorkDirExp = @"\\chenas03\Home$\40385$\";
            //strCeninwDirExpSupport = "";
            //strNASExplorer = "";
            //strUniqueId = "KW115-K13662";

            if (strEmpLogin == "")
            {
                strEmpLogin = "None";
            }
            if (strNASExplorer != "")
            {
                strNASExplorer = strNASExplorer.ToLower();
                if (strCeninwExp.Contains("chenas03") || strCenproExp.Contains("chenas03"))
                {
                    strNASExplorer = "";
                }                
            }
            if (strCommand != "")
                strCommand = objDS.EncryptData(strCommand);
            //if (strCeninwDirExpSupport != "")
            //    strCeninwDirExpSupport = objDS.EncryptData(strCeninwDirExpSupport.Replace("[***]", "-"));
            //if (strCeninwExp != "")
            //    strCeninwExp = objDS.EncryptData(strCeninwExp.Replace("[***]", "-"));
            //if (strCenproExp != "")
            //    strCenproExp = objDS.EncryptData(strCenproExp.Replace("[***]", "-"));
            //if (strCenAppExp != "")
            //    strCenAppExp = objDS.EncryptData(strCenAppExp.Replace("[***]", "-"));
            //if (strWorkDirExp != "")
            //    strWorkDirExp = objDS.EncryptData(strWorkDirExp.Replace("[***]", "-"));
            // Validate URL and send mail
            string strPasswordRequired = "No";
            string strCurrentURL = Request.Url.ToString().ToLower();
            string strPassParam = string.Empty;
            if (strPasswordRequired != "")
                strPasswordRequired = objDS.EncryptData(strPasswordRequired);

            if (strNASExplorer == "withcenapp")
                strNASExplorer = "NASExplorerNew";
            else
                strNASExplorer = "NASExplorer";

            string url = string.Empty;
            if (strWorkDirExp == "" && strCeninwExp != "" && (strCenproExp != "" || strCenAppExp != ""))
                url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Ceninw=" + strCeninwExp + "&Cenpro=" + strCenproExp + "&CenApp=" + strCenAppExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            else if (strWorkDirExp != "" && strCeninwExp == "" && strCenproExp == "" && strCenAppExp == "")
                url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&WorkDir=" + strWorkDirExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            else if (strWorkDirExp != "" && strCeninwExp != "" && strCenproExp == "" && strCenAppExp == "")
                url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Ceninw=" + strCeninwExp + "&WorkDir=" + strWorkDirExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            else if (strWorkDirExp != "" && strCeninwExp == "" && (strCenproExp != "" || strCenAppExp != ""))
                url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Cenpro=" + strCenproExp + "&CenApp=" + strCenAppExp + "&WorkDir=" + strWorkDirExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            else if (strWorkDirExp == "" && strCeninwExp == "" && (strCenproExp != "" || strCenAppExp != ""))
                url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Cenpro=" + strCenproExp + "&CenApp=" + strCenAppExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            else if (strWorkDirExp == "" && strCeninwExp!= "" && strCenproExp == "" && strCenAppExp == "" )
            url = "../" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Ceninw=" + strCeninwExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
        else if (strCeninwDirExpSupport != "" && strWorkDirExp == "" && strCeninwExp == "" && strCenproExp == "" && strCenAppExp == "" )
           strCommand = objDS.DecryptData(strCommand);
            if (strCommand.Contains("SUPPORT-RIGHTS"))
            {
                if (strCommand.Contains("ROOTACCESS"))
                    strCommand = "ROOTACCESS|SUPPORT-RIGHTS|CINF|CIRE|CIUP|CICP|CIPS|CIDL";
                else
                    strCommand = "SUPPORT-RIGHTS|CINF|CIRE|CIUP|CICP|CIPS|CIDL";

                strCommand = objDS.EncryptData(strCommand);  //https://smarttrack.kwglobal.com/smarttrack-ch/
                url = "https://smarttrack.kwglobal.com/smarttrack-ch/" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Ceninw=" + strCeninwDirExpSupport + "&IsSupport=Yes&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "";
                //url = "http://localhost:44358/NasExplorer/XPlorer";// + strCommand + "&Ceninw=" + strCeninwDirExpSupport + "&IsSupport=Yes&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "";
            }
            else
                url = "https://smarttrack.kwglobal.com/smarttrack-ch/" + strNASExplorer + ".aspx?strCommand=" + strCommand + "&Ceninw=" + strCeninwExp + "&Cenpro=" + strCenproExp + "&CenApp=" + strCenAppExp + "&WorkDir=" + strWorkDirExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
                //url = "http://localhost:44358/NasExplorer/XPlorer";// + strCommand + "&Ceninw=" + strCeninwExp + "&Cenpro=" + strCenproExp + "&CenApp=" + strCenAppExp + "&WorkDir=" + strWorkDirExp + "&empId=" + strEmpLogin + "&Required=" + strPasswordRequired + strPassParam + "&uniqueId=" + strUniqueId;
            string s = "window.open('" + url + "', 'popup_window', 'width=" + (intW + 300).ToString() + "px,height=" + (intH - 300).ToString() + "px,left=0,top=0,resizable=yes','toolbar=no','location=no','addressbar=no');";

            return Json(url, JsonRequestBehavior.AllowGet);
     
            // ScriptManager.RegisterClientScriptBlock(Me.Page, Me.GetType(), "script", s, True)

            //  return View("../NasExplorer/NasExplorer");
        }       
        
       public ActionResult NASExplorer()
        {
            return View();
        }
        public ActionResult XPlorer()
        {
            return View();
        }
        public PartialViewResult getExplorerWnd()
        {
            NasExpModel model = new NasExpModel();
            return PartialView("NASDirectory", model);
            
        }

        [SessionExpire]
        public ActionResult GetList()
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select top 50 EmpLogin,EmpName,DeptCode,CustAccess from JBM_EmployeeMaster", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}