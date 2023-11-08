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
using ReferenceLibrary.IniDet;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using System.Security.Principal;
using ExcelDataReader;
using System.Text;

namespace SmartTrack.Controllers
{
    public class CleanUpController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: CleanUp
        public ActionResult Index()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult CleanUpInfo()
        {
            try
            {
                string strquery = "select SeqNo,FolderPath,FileType,NoOfDays,IncludeSubDirectory,TopDirectoryOnly,DeleteSubDir,DeleteTopDir,Project from JBM_CleanUp";
                DataTable dtart = new DataTable();
                dtart = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());               

                DataSet ds = new DataSet();
                ds.Tables.Add(dtart);
                return View(ds);
            }
            catch {
                return View();
            }
        }
        [SessionExpire]
        public ActionResult AddNewFolderDetails(string sFolderName, string sFileType,string sNoOfDays,string sIncludeSubDirectory,string sTopDirectoryOnly,string sDeleteSubDir,string sDeleteTopDir,string sProject)
        {
            try
            {
                if (sFolderName.ToString() != "")
                {
                    DataTable dt = DBProc.GetResultasDataTbl("Select FolderPath from JBM_CleanUp Where FolderPath='" + sFolderName + "' ", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                    }
                    string strQuery = "INSERT INTO JBM_CleanUp (FolderPath,FileType,NoOfDays,IncludeSubDirectory,TopDirectoryOnly,DeleteSubDir,DeleteTopDir,Project) VALUES ('" + sFolderName + "', '" + sFileType + "', '" + sNoOfDays + "', '" + sIncludeSubDirectory + "', '" + sTopDirectoryOnly + "','" + sDeleteSubDir + "','" + sDeleteTopDir + "','" + sProject + "')";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

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
        [SessionExpire]
        public ActionResult UpdateCD(string sArtid, string sFolderName, string sFileType, string sNoOfDays, string sIncludeSubDirectory, string sTopDirectoryOnly, string sDeleteSubDir, string sDeleteTopDir, string sProject)
        {
            try
            {
                if (sArtid.ToString() != "")
                {

                    string strQuery = "UPDATE JBM_CleanUp  SET FolderPath = '" + sFolderName + "', FileType = '" + sFileType + "', NoOfDays= '" + sNoOfDays + "', IncludeSubDirectory = '" + sIncludeSubDirectory + "', TopDirectoryOnly = '" + sTopDirectoryOnly + "', DeleteTopDir = '" + sDeleteTopDir + "', DeleteSubDir = '" + sDeleteSubDir + "', Project = '" + sProject + "' WHERE [SeqNo] = '" + sArtid + "'";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());


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

        [SessionExpire]
        public ActionResult DeleteCD(string sArtid)
        {
            try
            {
                if (sArtid.ToString() != "")
                {                   
                    string strQuery = "DELETE FROM JBM_CleanUp WHERE SeqNo='" + sArtid + "'";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());


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