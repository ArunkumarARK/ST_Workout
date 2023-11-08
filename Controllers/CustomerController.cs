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
using System.Runtime.InteropServices;
using System.Reflection;

namespace SmartTrack.Controllers
{
    public class CustomerController : Controller
    {
        //[DllImport("kernel32", SetLastError = true)]
        // private static extern int GetPrivateProfileString(string sectionName,string keyName,string defaultValue,byte[] iniValue,int bufferLen,string filename);
        //private static extern int GetPrivateProfileSection(string lpAppName, string lpReturnedString, int nSize, string lpFileName);
        [DllImport("Kernel32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        private static extern UInt32 GetPrivateProfileSection
           (
               [In] [MarshalAs(UnmanagedType.LPStr)] string strSectionName,
               // Note that because the key/value pars are returned as null-terminated
               // strings with the last string followed by 2 null-characters, we cannot
               // use StringBuilder.
               [In] IntPtr pReturnedString,
               [In] UInt32 nSize,
               [In] [MarshalAs(UnmanagedType.LPStr)] string strFileName
           );

        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: Customer
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Project()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult AddCustomer()
        {
            try
            {
                Session["PrjTabHead"] = "Customer Information";
                Session["PageHead"] = "Customer Information";
                if (Session["sesStatus"] == null)
                {
                    Session["PageSubHead"] = "Add New Customer";
                }
                else { Session["PageSubHead"] = "Update Customer"; }
                //Load Approve List
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                dsPM = DBProc.GetResultasDataSet("select EmpAutoID,EmpName from JBM_employeemaster where roleid='103'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsPM.Tables[0].Rows[intCount]["EmpAutoID"].ToString();
                        string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                        lstPM.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.PMlist = lstPM;

                //Load Category
                List<SelectListItem> lstStage = new List<SelectListItem>();
                DataSet dsStage = new DataSet();
                string strsquery = " select distinct Stage from JBM_FTP_Details where Stage is not null and Stage!=''";
                
                dsStage = DBProc.GetResultasDataSet(strsquery, Session["sConnSiteDB"].ToString());
                if (dsStage.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsStage.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsStage.Tables[0].Rows[intCount]["Stage"].ToString();
                        string strEmpName = dsStage.Tables[0].Rows[intCount]["Stage"].ToString();
                        lstStage.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Stagelist = lstStage;

                //Load Category
                List<SelectListItem> lstCategory= new List<SelectListItem>();
                DataSet dsCategory = new DataSet();
                string strquery= "Select distinct(custCategory) FROM JBM_CustomerMaster ";
                //if(Session["sCustAcc"]!=null)
                //    strquery+= "where CustType = '"+ Session["sCustAcc"].ToString()+ "'";
                dsCategory = DBProc.GetResultasDataSet(strquery, Session["sConnSiteDB"].ToString());
                if (dsCategory.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCategory.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsCategory.Tables[0].Rows[intCount]["custCategory"].ToString();
                        string strEmpName = dsCategory.Tables[0].Rows[intCount]["custCategory"].ToString();
                        lstCategory.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Categorylist = lstCategory;

                //Load Team
                List<SelectListItem> lstTeam = new List<SelectListItem>();
                DataSet dsTeam = new DataSet();
                strquery = "select TeamID, Description from JBM_CustTeamID   ";
                if (Session["sCustAcc"] != null)
                    strquery += "where custType='" + Session["sCustAcc"].ToString() + "'";
                strquery += " order by teamid desc";
                dsTeam = DBProc.GetResultasDataSet(strquery, Session["sConnSiteDB"].ToString());
                if (dsTeam.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsTeam.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsTeam.Tables[0].Rows[intCount]["TeamID"].ToString();
                        string strEmpName = dsTeam.Tables[0].Rows[intCount]["Description"].ToString();
                        lstTeam.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Teamlist = lstTeam;

                //Load ROOT
                List<SelectListItem> lstRoot= new List<SelectListItem>();
                DataSet dsRoot = new DataSet();
                dsRoot = DBProc.GetResultasDataSet("Select RootPath, RootID from JBM_RootDirectory order by RootID", Session["sConnSiteDB"].ToString());
                if (dsRoot.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsRoot.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsRoot.Tables[0].Rows[intCount]["RootID"].ToString();
                        string strEmpName = dsRoot.Tables[0].Rows[intCount]["RootPath"].ToString();
                        lstRoot.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Rootlist = lstRoot;

                //Load Country
                List<SelectListItem> lstCountry = new List<SelectListItem>();
                DataSet dsCountry = new DataSet();
                dsCountry = DBProc.GetResultasDataSet("Select distinct(country) from JBM_CtyList ORDER BY Country desc", Session["sConnSiteDB"].ToString());
                if (dsCountry.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCountry.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsCountry.Tables[0].Rows[intCount]["country"].ToString();
                        string strEmpName = dsCountry.Tables[0].Rows[intCount]["country"].ToString();
                        lstCountry.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Countrylist = lstCountry;

                DataTable dtftp = new DataTable();
                dtftp.Columns.Add("ftpProfileName", typeof(String));
                dtftp.Columns.Add("ftpHost", typeof(String));
                dtftp.Columns.Add("ftpUID", typeof(String));
                dtftp.Columns.Add("ftpPWD", typeof(String));
                dtftp.Columns.Add("ftpPath", typeof(String));                
                dtftp.Columns.Add("Stage", typeof(String));
                dtftp.Columns.Add("DestinationPath", typeof(String));
                dtftp.Columns.Add("CutorCopy", typeof(String));
                dtftp.Columns.Add("FTPID", typeof(String));
                dtftp.Columns.Add("CustAccess", typeof(String));
                Session.Add("dtftp", dtftp);

                Session["sStatus"] = null;
                if (Session["sesCustID"].ToString()!="" && Session["sesCustID"].ToString() != null)
                {
                    if (Session["sesStatus"].ToString().Trim() != "" && Session["sesStatus"].ToString().Trim() != null)
                    {
                        DataTable ds = new DataTable();
                        string strQueryFinal = "SELECT  jc.CustID,jc.CustName, jc.CustCategory, jc.CustType, jc.CustSN, jt.Description, jc.JBM_TeamID, jc.CustAddress1,jc.CustAddress2, jc.CustCity, jc.CustState, jc.CustCountry, jc.CustPhone, jc.CustFax,jc.CustEmail, jr.RootPath, jr.RootID, jc.CustWebAddress, jr1.RootPath as RootPathCeninw, jc.RootIDCeninw, jf.FTPID,jf.ftpProfileName, jf.ftpHost, jf.ftpUID, jf.ftpPWD, jf.ftpPath ,(case when jc.Cust_Disabled = '0' then 'Enabled' when jc.Cust_Disabled = '1' then 'Disabled' End) as Cust_Disabled,jf.CustAccess,jf.Stage,jf.DestinationPath,jf.CutOrCopy,jc.JBM_Iss_UnAssigned FROM JBM_CustomerMaster jc left JOIN JBM_Ftp_Details jf ON jc.CustID = jf.CustID inner join JBM_CustTeamID jt on jc.JBM_TeamID = jt.TeamID inner join JBM_RootDirectory jr on jc.RootID = jr.RootID inner join JBM_RootDirectory jr1 on jc.RootIDCeninw = jr1.RootID where jc.CustID = '" + Session["sesCustID"].ToString().Trim() + "'";
                        ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                        if (ds.Rows.Count > 0)
                        {
                            ViewData.Model = ds.AsEnumerable();
                            Session["sStatus"] = Session["sesStatus"].ToString();
                        }
                    }
                }
                Session["sesCustID"] = null;
                Session["sesStatus"] = null;
               
                return View();
            }
            catch {
                return View();
            }
        }
        [SessionExpire]
        public ActionResult GetCustomerPMList(string sCustomer,string sselcustname)
        {

            try
            {
                if (sCustomer != "" || sCustomer != null)
                {
                    if (Session["sStatus"] == null)
                    {
                        if (sCustomer.ToString() != "" && sselcustname.ToString() != "")
                        {
                            DataTable dt = DBProc.GetResultasDataTbl("Select CustCategory,CustName from JBM_CustomerMaster where CustCategory='" + sCustomer + "' and  CustName='" + sselcustname + "'", Session["sConnSiteDB"].ToString());
                            if (dt.Rows.Count > 0)
                            {
                                return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                            }
                        }
                    }
                    string strQueryFinal = "Select CustId,Custname from JBM_CustomerMaster where CustCategory='"+ sCustomer + "' and CustType='" + Session["sCustAcc"].ToString() + "' order by Custname desc";

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {a[0].ToString(),a[1].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetStateList(string sCustomer)
        {

            try
            {
                if (sCustomer != "" || sCustomer != null)
                {
                    string strQueryFinal = "Select State from JBM_CtyList WHERE Country='"+ sCustomer + "' order by State desc";

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {a[0].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetEmpMailID(string sEmpID,string sCustName,string sCustSN,string sCustType)
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpMailId,EmpName from JBM_EmployeeMaster where EMpAutoID='" + sEmpID + "'", Session["sConnSiteDB"].ToString());
                string EmpMailID = "";
                string ApproveEmpName = "";
                string MailBody = "";
                if (ds.Tables[0].Rows.Count > 0)
                {
                    EmpMailID = ds.Tables[0].Rows[0]["EmpMailId"].ToString();
                    ApproveEmpName = ds.Tables[0].Rows[0]["EmpName"].ToString();
                }
                else
                {
                    EmpMailID = Session["EmpMailId"].ToString();
                    ApproveEmpName = Session["EmpName"].ToString();
                }
                string strQueryFinal = "select MsgBody,Msgsubject from JBM_messageInfo where MsgID=108";
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ds.Tables[0].Rows[0]["MsgBody"] = ds.Tables[0].Rows[0]["MsgBody"].ToString().Replace("###ApprovalNAME###", ApproveEmpName).Replace("###EMPNAME###", Session["EmpName"].ToString()).Replace("###DATE###", DateTime.Now.ToString()).Replace("###CUSTNAME###", sCustName).Replace("###CUSTTYPE###", sCustType).Replace("###CUSTSN###", sCustSN);
                    MailBody = ds.Tables[0].Rows[0]["MsgBody"].ToString();
                }
                return Json(new { dataComp = EmpMailID , dataComps = MailBody }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetRootID()
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select RootPath, RootID from JBM_RootDirectory order by RootID", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString(),a[1].ToString()
                                };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetFolderIndex()
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select distinct FolderIndex from JBM_DeptFolders", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString()
                                };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetCityList(string sCountry,string sState)
        {

            try
            {
                if (sCountry != "" || sCountry != null)
                {
                    if (sState != "" || sState != null)
                    {
                        string strQueryFinal = "Select City from JBM_CtyList WHERE Country='" + sCountry + "' and State='"+ sState + "'";

                        DataSet ds = new DataSet();
                        ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                        var JSONString = from a in ds.Tables[0].AsEnumerable()
                                         select new[] {a[0].ToString()
                                };
                        return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                    }
                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetCustomerType(string sCustomerType,string sK,string sselCategory,string sselcustname,string sFirstchar)
        {
            try
            {
                if (sCustomerType != "" || sCustomerType != null)
                {
                    Session["Job_ID"] = sFirstchar;
                    if (Session["sStatus"] == null)
                    {
                        if (sselCategory.ToString() != "" && sselcustname.ToString() != "")
                        {
                            DataTable dt = DBProc.GetResultasDataTbl("Select CustCategory,CustName from JBM_CustomerMaster where CustCategory='" + sselCategory + "' and  CustName='" + sselcustname + "'", Session["sConnSiteDB"].ToString());
                            if (dt.Rows.Count > 0)
                            {
                                return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                            }
                        }
                        string[] CustomerType = sCustomerType.Split(',');
                        List<string> list = new List<string>(CustomerType);
                        List<string> listarray = new List<string>();

                        String[] str = listarray.ToArray();
                        if (sK == "2")
                        {

                            for (int i = 0; i < list.Count; i++)
                            {
                                if (list[i][0] == list[i][1])
                                {
                                    listarray.Add(list[i]);
                                }
                            }
                            for (int i = 0; i < listarray.Count; i++)
                            {
                                list.Remove(listarray[i]);
                            }
                            string strQueryFinal = "select CustType from JBM_CustTeamID";

                            DataTable ds = new DataTable();
                            ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                            List<DataRow> listdt = ds.AsEnumerable().ToList();
                            for (int j = 0; j < list.Count; j++)
                            {
                                DataRow[] foundCT = ds.Select("CustType = '" + list[j] + "'");
                                if (foundCT.Length == 0)
                                {
                                    return Json(new { dataComp = list[j] }, JsonRequestBehavior.AllowGet);
                                }
                            }
                        }
                        else
                        {

                            for (int i = 0; i < list.Count; i++)
                            {
                                if (list[i].Length == 2)
                                {
                                    listarray.Add(list[i]);
                                }
                                else if (list[i][0] == list[i][1]) { listarray.Add(list[i]); }
                                else if (list[i][1] == list[i][2]) { listarray.Add(list[i]); }
                                else if (list[i][2] == list[i][3]) { listarray.Add(list[i]); }
                                else if (list[i][0] == list[i][1] && list[i][0] == list[i][2])
                                {
                                    listarray.Add(list[i]);
                                }
                                else if (list[i][0] == list[i][2] && list[i][1] == list[i][3])
                                {
                                    listarray.Add(list[i]);
                                }
                            }
                            for (int i = 0; i < listarray.Count; i++)
                            {
                                list.Remove(listarray[i]);
                            }
                            string strQueryFinal = "select CustSN from JBM_CustomerMaster";

                            DataTable ds = new DataTable();
                            ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                            List<DataRow> listdt = ds.AsEnumerable().ToList();
                            for (int j = 0; j < list.Count; j++)
                            {
                                DataRow[] foundCT = ds.Select("CustSN = '" + list[j] + "'");
                                if (foundCT.Length == 0)
                                {
                                    return Json(new { dataComp = list[j] }, JsonRequestBehavior.AllowGet);
                                }
                            }
                        }
                    }

                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetRootFolder(string sCustomerSN)
        {
            try
            {
                if (sCustomerSN.ToString() != "")
                {                   
                    //Load Root Folder
                    List<SelectListItem> lstDF = new List<SelectListItem>();
                    DataSet dsDF = new DataSet();
                    if (Session["sStatus"] != null)
                    {
                        if (Session["sSiteLocation"].ToString() == "Chennai")
                            dsDF = DBProc.GetResultasDataSet(" select RootID,RootPath from jbm_rootdirectory where  RootPath like '%"+ sCustomerSN + "%'", Session["sConnSiteDB"].ToString());
                        else
                            dsDF = DBProc.GetResultasDataSet(" select RootID,RootPath from jbm_rootdirectory where  RootPath like '%"+ sCustomerSN + "%' and RootPath  not like '%kglpropdf%'", Session["sConnSiteDB"].ToString());
                    }
                    else
                    {
                        if (Session["sSiteLocation"].ToString() == "Chennai")
                            dsDF = DBProc.GetResultasDataSet(" select RootID,RootPath from jbm_rootdirectory where  RootPath like '%AIP%'", Session["sConnSiteDB"].ToString());
                        else
                            dsDF = DBProc.GetResultasDataSet(" select RootID,RootPath from jbm_rootdirectory where  RootPath like '%AIP%' and RootPath  not like '%kglpropdf%'", Session["sConnSiteDB"].ToString());
                        if (dsDF.Tables[0].Rows.Count > 0)
                        {
                            for (int intCount = 0; intCount < dsDF.Tables[0].Rows.Count; intCount++)
                            {
                                dsDF.Tables[0].Rows[intCount]["RootPath"] = dsDF.Tables[0].Rows[intCount]["RootPath"].ToString().Replace("AIP", sCustomerSN);
                            }

                        }
                    }
                    
                    var JSONString = from a in dsDF.Tables[0].AsEnumerable()
                                     select new[] {a[0].ToString(),a[1].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult Savedeptfolder(string sfolderpath,string scusttype,string sRootID, string sProPdfDir, string sFolderIndex)
        {
            try
            {
                if (sfolderpath.ToString() != "")
                {
                    DataTable dt = DBProc.GetResultasDataTbl("select * from JBM_DeptFolders where CustType='" + scusttype + "' and FolderIndex= '" + sFolderIndex + "'", Session["sConnSiteDB"].ToString());
                    if(dt.Rows.Count>0)
                    {
                        return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                    }
                    string strQuery = "Insert into JBM_DeptFolders (CustType,RootID,PropdfDir,FolderDir,FolderIndex,CreateDirectory) values('"+ scusttype + "','"+ sRootID + "','"+ sProPdfDir + "','" + sfolderpath + "','"+ sFolderIndex + "',0)";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    //Load Dept Folder
                    List<SelectListItem> lstDF = new List<SelectListItem>();
                    DataSet dsDF = new DataSet();
                    dsDF = DBProc.GetResultasDataSet("select '" + scusttype + "' as CustType,'0' as RootID,'" + sfolderpath + "' as FolderDir", Session["sConnSiteDB"].ToString());
                    
                    var JSONString = from a in dsDF.Tables[0].AsEnumerable()
                                     select new[] {a[1].ToString(),a[2].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
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
        public ActionResult Deletedeptfolder(string sfolderpath, string scusttype)
        {
            try
            {
                if (sfolderpath.ToString() != "")
                {
                    string strQuery = "Delete from JBM_DeptFolders where CustType='" + scusttype + "' and FolderDir='" + sfolderpath + "'";
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
        public ActionResult DeleteFTPDetails(string sftpid)
        {
            try
            {
                if (sftpid.ToString() != "")
                {
                    if (Session["sStatus"] == null)
                    {
                        DataTable ds = Session["dtftp"] as DataTable;
                        //ds.Select("FTPID = " + sftpid);
                        for (int i = ds.Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow dr = ds.Rows[i];
                            if (dr["FTPID"].ToString() == sftpid)
                                dr.Delete();
                        }
                        ds.AcceptChanges();
                        Session.Add("dtftp", ds);
                    }
                    else
                    {
                        string strQuery = "Delete from JBM_FTP_Details where FTPID='" + sftpid + "'";
                        string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
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
        [SessionExpire]
        public ActionResult GetFTPDetails(string sftpid)
        {
            try
            {
                if (sftpid.ToString() != "")
                {
                    DataTable dt = new DataTable();
                    if (Session["sStatus"] == null)
                    {
                        dt.Columns.Add("ftpProfileName", typeof(String));
                        dt.Columns.Add("ftpHost", typeof(String));
                        dt.Columns.Add("ftpUID", typeof(String));
                        dt.Columns.Add("ftpPWD", typeof(String));
                        dt.Columns.Add("ftpPath", typeof(String));
                        dt.Columns.Add("Stage", typeof(String));
                        dt.Columns.Add("DestinationPath", typeof(String));
                        dt.Columns.Add("CutorCopy", typeof(String));
                        dt.Columns.Add("FTPID", typeof(String));
                        dt.Columns.Add("CustAccess", typeof(String));
                        DataTable ds = Session["dtftp"] as DataTable;
                        //ds.Select("FTPID = "+ sftpid);
                        foreach (DataRow dr in ds.Rows)
                        {
                            if (dr["FTPID"].ToString() == sftpid)
                                dt.Rows.Add(dr.ItemArray);
                        }
                    }
                    else
                    {
                        string strQuery = "Select ftpProfileName,ftpHost,ftpUID,ftpPWD,ftpPath,Stage,DestinationPath,CutOrCopy from JBM_FTP_Details where FTPID='" + sftpid + "'";
                        dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                    }
                    
                    var JSONString = from a in dt.AsEnumerable()
                                     select new[] {a[0].ToString(),a[1].ToString(),a[2].ToString(),a[3].ToString(),a[4].ToString(),a[5].ToString(),a[6].ToString(),a[7].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
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
        public ActionResult Editdeptfolder(string sfolderpath, string scusttype)
        {
            try
            {
                if (sfolderpath.ToString() != "")
                {
                     string strQuery = "select * from JBM_DeptFolders where CustType='"+ scusttype + "' and FolderDir='"+ sfolderpath + "'";
                    //string strQuery = " select * from JBM_DeptFolders where CustType='AM' and FolderDir='Support\\Production\\APMA\\Download\\AdditionalCorrections\\'";
                   
                    DataSet dsDF = new DataSet();
                    dsDF = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in dsDF.Tables[0].AsEnumerable()
                                     select new[] {a[1].ToString(),a[2].ToString(),a[3].ToString(),a[4].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult Updatedeptfolder(string snewfolderpath, string scusttype, string sRootID, string sProPdfDir, string sFolderIndex, string sfolderpath)
        {
            try
            {
                if (sfolderpath.ToString() != "")
                {
                    string strQuery = "update JBM_DeptFolders set RootID='"+ sRootID + "',PropdfDir='"+ sProPdfDir + "',FolderDir='"+ snewfolderpath + "',FolderIndex='"+ sFolderIndex + "',CreateDirectory='0' where CustType='" + scusttype + "' and FolderDir='" + sfolderpath + "'";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                   
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult saverootfolder(string sCustomerSN, string scusttype, string sRootlst)
        {
            try
            {
                if (sCustomerSN.ToString() != "" && sRootlst.ToString() != "")
                {
                    string[] RootIds = sRootlst.Split(',');
                    List<string> list = new List<string>(RootIds);

                    string CenproRootID = "";
                    string CeninwRootID = "";
                    string ProPdfDirID = "";
                    List<SelectListItem> lstDF = new List<SelectListItem>();
                    DataSet dsDF = new DataSet();
                    if (Session["sStatus"] == null)
                    {
                        // Folder Creation
                        for (int count = 0; count < list.Count; count++)
                        {
                            string Path = @"" + list[count].ToString();

                            if (!Directory.Exists(Path))
                            {
                                Directory.CreateDirectory(Path);
                            }
                        }


                        DataTable dt = new DataTable();

                        dt = DBProc.GetResultasDataTbl("select max(CAST(RootID AS int)+1) as RootID from JBM_RootDirectory  ", Session["sConnSiteDB"].ToString());
                        CeninwRootID = dt.Rows[0][0].ToString();
                        Session["CeninwRootID"] = CeninwRootID;
                        string strQuery = "Insert into JBM_RootDirectory(RootID, RootPath) values ('" + CeninwRootID + "','" + list[0] + "')";
                        string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        dt = DBProc.GetResultasDataTbl("select max(CAST(RootID AS int)+1) as RootID from JBM_RootDirectory  ", Session["sConnSiteDB"].ToString());
                        CenproRootID = dt.Rows[0][0].ToString();
                        Session["CenproRootID"] = CenproRootID;
                        strQuery = "Insert into JBM_RootDirectory(RootID, RootPath) values ('" + CenproRootID + "','" + list[1] + "')";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        //Update RootID in customer master and account desc tables

                        strQuery = "update JBM_CustomerMaster set RootID='" + CenproRootID + "',RootIDCeninw='" + CeninwRootID + "' where CustType='" + scusttype + "'";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                        strQuery = "update JBM_AccountTypeDesc set RootID='" + CenproRootID + "',InwardDir='" + CeninwRootID + "' where CustAccess='" + scusttype + "'";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        if (Session["sSiteLocation"].ToString() == "Chennai")
                        {
                            dt = DBProc.GetResultasDataTbl("select max(CAST(RootID AS int)+1) as RootID from JBM_RootDirectory  ", Session["sConnSiteDB"].ToString());
                            ProPdfDirID = dt.Rows[0][0].ToString();
                            Session["ProPdfDirID"] = ProPdfDirID;
                            strQuery = "Insert into JBM_RootDirectory(RootID, RootPath) values ('" + ProPdfDirID + "','" + list[2] + "')";
                            strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                            strQuery = "update JBM_AccountTypeDesc set ProPdfDir='" + ProPdfDirID + "'where CustAccess='" + scusttype + "'";
                            strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                        }
                        
                    }
                    
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetDeptFolder(string sCustomerSN,string scusttype,string sRootlst)
        {
            try
            {
                if (sCustomerSN.ToString() != "" && sRootlst.ToString() != "")
                {    
                    DataSet dsDF = new DataSet();
                    if (Session["sStatus"] == null)
                    {
                        //Load Dept Folder
                       
                        dsDF = DBProc.GetResultasDataSet("select CustType ,RootID ,FolderDir,FolderIndex ,CreateDirectory from JBM_DeptFolders where CustType='CS'", Session["sConnSiteDB"].ToString());
                        if (dsDF.Tables[0].Rows.Count > 0)
                        {
                            for (int intCount = 0; intCount < dsDF.Tables[0].Rows.Count; intCount++)
                            {
                                dsDF.Tables[0].Rows[intCount]["FolderDir"] = dsDF.Tables[0].Rows[intCount]["FolderDir"].ToString().Replace("CSP", sCustomerSN);

                            }

                        }
                    }
                    else
                    {
                        dsDF = DBProc.GetResultasDataSet("select CustType ,RootID ,FolderDir,FolderIndex ,CreateDirectory from JBM_DeptFolders where CustType='"+ scusttype + "'", Session["sConnSiteDB"].ToString());
                    }
                    var JSONString = from a in dsDF.Tables[0].AsEnumerable()
                                     select new[] {a[1].ToString(),a[2].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }
                else

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetFolderlist(string sCustomerSN, string scusttype)
        {
            try
            {
                if (sCustomerSN.ToString() != "" )
                {
                    //Load Folder
                    List<string> lines = System.IO.File.ReadLines("D:\\list.txt").ToList();
                    DataTable dt = new DataTable();
                    dt.Columns.Add("FolderName", typeof(String));

                    for(int i=0;i<lines.Count;i++)
                    {
                        lines[i] = lines[i].ToString().Replace("###CUSTSN###", sCustomerSN);
                        dt.Rows.Add(lines[i]);
                    }

                    var JSONString = from a in dt.AsEnumerable()
                                     select new[] {a[0].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult Foldercreation(string sfolderlist, string scusttype)
        {
            try
            {
                if (sfolderlist.ToString() != "")
                {
                    string[] RootIds = sfolderlist.Split('\n');
                    string[] lines = sfolderlist.Split(new string[] { "\r\n", "\r", "\n" },StringSplitOptions.None);
                    List<string> list = new List<string>(RootIds);

                    // Folder Creation
                    for (int count = 0; count < list.Count; count++)
                    {
                        string Path = @"" + list[count].ToString();

                        if (!Directory.Exists(Path))
                        {
                            Directory.CreateDirectory(Path);
                        }
                    }
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult Folderupdation(string sfolderlist, string scusttype, string sCustomerSN)
        {
            try
            {
                if (sfolderlist.ToString() != "")
                {
                    FileInfo file = new FileInfo("D:\\list.txt");
                    System.IO.File.WriteAllText("D:\\list.txt", String.Empty);

                    string[] RootIds = sfolderlist.Split(',');
                    List<string> list = new List<string>(RootIds);

                    
                    for (int count = 0; count < list.Count; count++)
                    {
                        using (StreamWriter sw = file.AppendText())
                        {
                            list[count] = list[count].ToString().Replace(sCustomerSN,"###CUSTSN###");
                            sw.WriteLine(list[count].ToString());
                        }
                    }
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult DeptSave(string scusttype, string scustsn,string sRootlst)
        {
            try
            {
                if (scusttype.ToString() != "" && sRootlst.ToString() != "" && scustsn.ToString()!="")
                {
                    string[] RootIds = sRootlst.Split(',');
                    List<string> list = new List<string>(RootIds);
                    DataSet dsDF = new DataSet();
                   
                    if (Session["sStatus"] == null)
                    {
                        string CenproRootID = Session["CenproRootID"].ToString();
                        string CeninwRootID = Session["CeninwRootID"].ToString();
                        string ProPdfDirID = "";
                        if (Session["ProPdfDirID"] != null)
                            ProPdfDirID = Session["ProPdfDirID"].ToString();
                        string strQuery = "Insert into JBM_DeptFolders (CustType,RootID,FolderDir,FolderIndex,CreateDirectory) select '" + scusttype + "' as CustType,'" + CeninwRootID + "' as RootID,FolderDir,FolderIndex,CreateDirectory from JBM_DeptFolders where CustType ='CS' and RootID='74'";
                        string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        strQuery = "Insert into JBM_DeptFolders (CustType,RootID,FolderDir,FolderIndex,CreateDirectory) select '" + scusttype + "' as CustType,'" + CenproRootID + "' as RootID,FolderDir,FolderIndex,CreateDirectory from JBM_DeptFolders where CustType ='CS' and RootID='75'";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        strQuery = "Insert into JBM_DeptFolders (CustType,RootID,FolderDir,FolderIndex,CreateDirectory) select '" + scusttype + "' as CustType, RootID,FolderDir,FolderIndex,CreateDirectory from JBM_DeptFolders where CustType ='CS' and RootID='5'";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        strQuery = "Insert into JBM_DeptFolders (CustType,RootID,FolderDir,FolderIndex,CreateDirectory) select '" + scusttype + "' as CustType, RootID,FolderDir,FolderIndex,CreateDirectory from JBM_DeptFolders where CustType ='CS' and RootID='3'";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                        strQuery = "Insert into JBM_DeptFolders (CustType,RootID,FolderDir,FolderIndex,CreateDirectory) select '" + scusttype + "' as CustType, RootID,FolderDir=Replace(FolderDir,'\\CSP\\','\\" + scustsn + "\\'),FolderIndex,CreateDirectory from JBM_DeptFolders where CustType ='CS' and RootID='4';";
                        strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                                               
                        dsDF = DBProc.GetResultasDataSet("select * from JBM_DeptFolders where CustType='" + scusttype + "' and CreateDirectory=0", Session["sConnSiteDB"].ToString());
                        if (dsDF.Tables[0].Rows.Count > 0)
                        {
                            for (int intCount = 0; intCount < dsDF.Tables[0].Rows.Count; intCount++)
                            {
                                //Folder Creation 
                                string RootID = dsDF.Tables[0].Rows[intCount]["RootID"].ToString();
                                DataTable dtrootid = new DataTable();
                                dtrootid = DBProc.GetResultasDataTbl("select * from JBM_RootDirectory where RootID='" + RootID + "' ", Session["sConnSiteDB"].ToString());

                                string Rootpath = @"" + dtrootid.Rows[0]["RootPath"].ToString() + dsDF.Tables[0].Rows[intCount]["FolderDir"].ToString();

                                if (!Directory.Exists(Rootpath))
                                {
                                    Directory.CreateDirectory(Rootpath);
                                }
                            }

                        }
                    }
                   
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else

                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetFTPList(string scusttype)
        {
            try
            {
                DataTable  ds = new DataTable();
                if (scusttype.ToString() != "")
                {
                    if (Session["sStatus"] == null)
                    {
                        ds = Session["dtftp"] as DataTable;
                    }
                    else
                    {
                        string strQueryFinal = "select ftpProfileName,ftpHost,ftpUID,ftpPWD,ftpPath,Stage,DestinationPath,CutOrCopy,FTPID from JBM_FTP_Details where CustAccess='" + scusttype + "'";
                        ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                    }
                }
               
                var JSONString = from a in ds.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString(),
                                     CreateBtn(a[8].ToString())
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult Getmaildetails(string scusttype, string sCustSN,string sCustName)
        {
            try
            {
                DataSet ds = new DataSet();
                if (scusttype.ToString() != "")
                {
                    if (Session["sStatus"] == null)
                    {
                        string strQueryFinal = "select MsgBody,Msgsubject from JBM_messageInfo where MsgID=108";
                        ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            ds.Tables[0].Rows[0]["Msgsubject"] = ds.Tables[0].Rows[0]["Msgsubject"].ToString().Replace("###CUSTNAME###", sCustName);
                            ds.Tables[0].Rows[0]["MsgBody"] = ds.Tables[0].Rows[0]["MsgBody"].ToString().Replace("###EMPNAME###", Session["EmpName"].ToString()).Replace("###DATE###", DateTime.Now.ToString()).Replace("###CUSTNAME###", sCustName).Replace("###CUSTTYPE###", scusttype).Replace("###CUSTSN###", sCustSN);
                        }
                    }
                }
                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult FTPSave(string sProfileName, string sFtphost, string sFtppath, string sUserid, string sPwd, string sCustAccess, string sStage, string sDestinationPath, string sCutorCopy,string sCustID, string sftpid)
        {
            try
            {
                if (sftpid == "")
                {
                    string CustID = "";
                    if (Session["sStatus"] == null)
                        CustID = Session["FTPCustID"].ToString();
                    else
                        CustID = sCustID;

                    DataTable dtftp = Session["dtftp"] as DataTable;
                    for (int i = 0; i < dtftp.Rows.Count; i++)
                    {
                        DataTable dt = DBProc.GetResultasDataTbl("Select CustId,FTPID, ftpProfileName, ftpHost, ftpUID, ftpPWD from JBM_Ftp_Details where CustID='" + CustID + "' and ftpProfileName='" + dtftp.Rows[i]["ftpProfileName"].ToString() + "'", Session["sConnSiteDB"].ToString());
                        if (dt.Rows.Count > 0)
                        {
                            string strCustomerid = dt.Rows[0]["CustId"].ToString();
                            string strFTPID = dt.Rows[0]["FTPID"].ToString();
                            string strProfilename = dt.Rows[0]["ftpProfileName"].ToString();
                            if (CustID == strCustomerid && sProfileName == strProfilename)
                            {
                                return Json(new { dataComp = strCustomerid + " already exists" }, JsonRequestBehavior.AllowGet);
                            }
                        }

                        string strQuery = "Insert into JBM_Ftp_Details(CustID, FTPID, ftpProfileName, ftpHost, ftpUID, ftpPWD, ftpPath,CustAccess,Stage,DestinationPath,CutorCopy) values ('" + CustID + "','" + dtftp.Rows[i]["FTPID"].ToString() + "','" + dtftp.Rows[i]["ftpProfileName"].ToString() + "','" + dtftp.Rows[i]["ftpHost"].ToString() + "','" + dtftp.Rows[i]["ftpUID"].ToString() + "','" + dtftp.Rows[i]["ftpPWD"].ToString() + "','" + dtftp.Rows[i]["ftpPath"].ToString() + "','" + dtftp.Rows[i]["CustAccess"].ToString() + "','" + dtftp.Rows[i]["Stage"].ToString() + "','" + dtftp.Rows[i]["DestinationPath"].ToString() + "','" + dtftp.Rows[i]["CutorCopy"].ToString() + "')";
                        string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                    }
                    
                }
                else
                {
                    DataTable dtftp = Session["dtftp"] as DataTable;
                    for (int i = 0; i < dtftp.Rows.Count; i++)
                    {
                        string strQuery = "update JBM_Ftp_Details set ftpProfileName='" + dtftp.Rows[i]["ftpProfileName"].ToString() + "',ftpHost='" + dtftp.Rows[i]["ftpHost"].ToString() + "',ftpUID='" + dtftp.Rows[i]["ftpUID"].ToString() + "',ftpPWD='" + dtftp.Rows[i]["ftpPWD"].ToString() + "',ftpPath='" + dtftp.Rows[i]["ftpPath"].ToString() + "',CustAccess='" + dtftp.Rows[i]["CustAccess"].ToString() + "',Stage='" + dtftp.Rows[i]["Stage"].ToString() + "',DestinationPath='" + dtftp.Rows[i]["DestinationPath"].ToString() + "',CutorCopy='" + dtftp.Rows[i]["CutorCopy"].ToString() + "' where FTPID='" + dtftp.Rows[i]["FTPID"].ToString() + "'";
                        string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                    }
                }


                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult FTPSavetodatatable(string sProfileName, string sFtphost, string sFtppath, string sUserid, string sPwd, string sCustAccess, string sStage, string sDestinationPath, string sCutorCopy, string sCustID, string sftpid)
        {
            try
            {
                if (sftpid == "")
                {
                    string strCustId = "";                   

                    DataTable dt = DBProc.GetResultasDataTbl("Select max(FTPID) as FTPID from JBM_Ftp_Details", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["FTPID"].ToString() != null || dt.Rows[0]["FTPID"].ToString().Trim() != "")
                        {
                            strCustId = dt.Rows[0]["FTPID"].ToString();
                            strCustId = (Convert.ToInt32(strCustId.Substring(1)) + 1).ToString();
                        }
                        switch (strCustId.Length)
                        {
                            case 1:
                                strCustId = "F00" + strCustId;
                                break;
                            case 2:
                                strCustId = "F0" + strCustId;
                                break;
                            default:
                                strCustId = "F" + strCustId;
                                break;
                        }
                    }
                    else
                    {
                        strCustId = "F001";
                    }
                    
                    DataTable dtftp = Session["dtftp"] as DataTable;
                    DataRow workRow = dtftp.NewRow();
                    workRow["ftpProfileName"] = sProfileName;
                    workRow["ftpHost"] = sFtphost;
                    workRow["ftpUID"] = sUserid;
                    workRow["ftpPWD"] = sPwd;
                    workRow["ftpPath"] = sFtppath;
                    workRow["CustAccess"] = sCustAccess;
                    workRow["Stage"] = sStage;
                    workRow["DestinationPath"] = sDestinationPath;
                    workRow["CutorCopy"] = sCutorCopy;
                    workRow["FTPID"] = strCustId;
                    dtftp.Rows.Add(workRow);
                }
                else
                {
                    DataTable dtftp = Session["dtftp"] as DataTable;
                    foreach (DataRow workRow in dtftp.Rows)
                    {
                        if (workRow["FTPID"].ToString() == sftpid)
                        {
                            workRow["ftpProfileName"] = sProfileName;
                            workRow["ftpHost"] = sFtphost;
                            workRow["ftpUID"] = sUserid;
                            workRow["ftpPWD"] = sPwd;
                            workRow["ftpPath"] = sFtppath;
                            workRow["CustAccess"] = sCustAccess;
                            workRow["Stage"] = sStage;
                            workRow["DestinationPath"] = sDestinationPath;
                            workRow["CutorCopy"] = sCutorCopy;
                        }
                    }
                }


                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }

        [SessionExpire]
        public ActionResult FinalApproval(string sApproveTo, string scustname, string scusttype, string sCustSN)
        {
            try
            {
                string strstart = "", strend = "";
                string strJM = "S1131";
                if (Session["sStatus"] == null)
                {
                    if (Session["RoleID"].ToString().Trim() == "103")
                    {
                        strstart = Init_Barcode.gstrProjectManagement;
                        strend = strJM;
                    }

                    string strArtStageTypeID = "";
                    string StrEmp = "";
                    string strStatus = "";
                    if (sApproveTo.Trim() == "0")
                    {
                        strArtStageTypeID = strstart + "_" + Init_Barcode.gstrArtPending + "_" + strend;
                        StrEmp = Session["EmpAutoId"].ToString();
                        strStatus = "Pending";
                    }
                    else
                    {
                        strArtStageTypeID = strstart + "_" + Init_Barcode.gstrArtCompleted + "_" + strend;
                        StrEmp = sApproveTo;
                        strStatus = "Pending";
                    }
                    string strWF_Code = "C1";

                    string strQuery = "Update JBM_CustomerMaster set AccountMgr='" + StrEmp + "',Status='" + strStatus + "',ArtStageTypeID='" + strArtStageTypeID + "',WF_Code='" + strWF_Code + "' where CustID='" + Session["FTPCustID"].ToString() + "'";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    strQuery = "update JBM_EmployeeMaster set TeamID=replace(REPLACE(TeamID,'" + Session["TempTeamID"].ToString() + "',''),'||','|')+'" + Session["TempTeamID"].ToString() + "|',CustAccess=(REPLACE(REPLACE(CustAccess,'" + scusttype + "',''),'||','|')+'" + scusttype + "|') where EMpAutoID='" + Session["EmpAutoId"].ToString() + "'";
                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    string strQueryFinal = "select MsgBody,Msgsubject from JBM_messageInfo where MsgID=110";
                    DataTable dtmail = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                    string mailSubject = "";
                    string mailBody = "";
                    if (dtmail.Rows.Count > 0)
                    {

                        mailSubject = dtmail.Rows[0]["Msgsubject"].ToString();
                        mailBody = dtmail.Rows[0]["MsgBody"].ToString().Replace("###CUSTTYPE###", scusttype);
                    }                    

                    strQuery = @"Insert into JBM_FTP_Stage_Details(CustType, JBM_AutoID, Stage, MailFrom, MailTo,
                MailCc, MailBcc, MailSubject, MailBody, MailErrorTo, MailErrorCc) values('" + scusttype + "','" + scusttype + "-Amail','Bookin','" + Session["EmpMailId"].ToString() + "','Sivagnanamoorthy@kwglobal.com; Mayavan.Renganathan@kwglobal.com; Nirmal.Kumar@kwglobal.com; thiyagarajan.g@kwglobal.com; aparna.MR@kwglobal.com; Sorimuthu.G@kwglobal.com',''," +
                   "'','110','110','MIS.Support@kwglobal.com','thiyagarajan.g@kwglobal.com')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    strQuery = @"Insert into JBM_MailInfo(EmpAutoID, JBM_AutoID, AutoArtID, MailEventID, MailStatus, 
                MailInitDate, Stage, MaxTry, First_try, MailFrom, MailTo, MailCc, MailBCc,
                MailSub, MailBody, IsBodyHtml, MailPriority, MailType, CustType) values('" + Session["EmpAutoId"] + "','','','E0001','1','" + DateTime.Now + "','','','','" + Session["EmpMailId"].ToString() + "'," +
                   ",'Sivagnanamoorthy@kwglobal.com; Mayavan.Renganathan@kwglobal.com; Nirmal.Kumar@kwglobal.com; thiyagarajan.g@kwglobal.com; aparna.MR@kwglobal.com; Sorimuthu.G@kwglobal.com','','','" + mailSubject + "','" + mailBody + "','1','2','Direct','" + scusttype + "')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());


                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult AddNewCustomer(string sCategory, string sKglteam, string sCenproPath, string sCeninwPath, string scustname, string scusttype, string sCustSN, string sCustID, string sCustomerStatus,string Iss)
        {
            try
            {             
                                                            
                string strCustId = "";
                string CustID = "";
                string TeamID = "";
                string strQuery = "";
                string strResult2 = "";
                DataTable dt = new DataTable(); 

                if (Session["sStatus"] == null)
                {
                    dt = DBProc.GetResultasDataTbl("select max(CAST(teamid AS int)+1) as teamid from JBM_CustTeamID ", Session["sConnSiteDB"].ToString());
                    TeamID = dt.Rows[0][0].ToString();
                    Session["TempTeamID"] = TeamID;
                    strQuery = "Insert into JBM_CustTeamID(CustType, TeamID, Description) values ('" + scusttype + "','" + TeamID + "','" + sCustSN + "')";
                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                    dt = DBProc.GetResultasDataTbl("Select max(CustID) as CustID from JBM_CustomerMaster", Session["sConnSiteDB"].ToString());

                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["CustID"].ToString() != null || dt.Rows[0]["CustID"].ToString().Trim() != "")
                        {
                            strCustId = dt.Rows[0]["CustID"].ToString();
                            strCustId = (Convert.ToInt32(strCustId.Substring(1)) + 1).ToString();
                        }
                        switch (strCustId.Length)
                        {
                            case 1:
                                strCustId = "C00" + strCustId;
                                break;
                            case 2:
                                strCustId = "C0" + strCustId;
                                break;
                            default:
                                strCustId = "C" + strCustId;
                                break;
                        }
                    }
                    else
                    {
                        strCustId = "C001";
                    }
                    dt = DBProc.GetResultasDataTbl("Select CustCategory, CustName from JBM_CustomerMaster where CustCategory='" + sCategory + "' and  CustName='" + scustname + "' and CustType='" + scusttype + "'", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        string strCustCateg = dt.Rows[0]["CustCategory"].ToString();
                        string strCustName = dt.Rows[0]["CustName"].ToString();
                        if (scustname == strCustName && sCategory == strCustCateg)
                        {
                            return Json(new { dataComp = "Category " + strCustCateg + " with Customer name " + strCustName + " already exists" }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    strQuery = @"Insert into JBM_AccountTypeDesc(CustAccess, AccountType, CustGroup, JBM_TeamID, Job_ID, Stages, 
                CompRptMenuItem, RevisesInfoMenuItem, ManualAllocationMenuItem, CreateDirectory, RootID, InwardDir, CustID, MLDir, 
                InwardFPDirPath, InwardRevDirPath, InwardIssueDirPath,    CeFpDirPath,   ReFpDirPath,   MLFpDirPath, 
                PagFpDirPath,PagIssueDirPath, PrFpDirPath,PrIssueDirPath, DispatchFpDirPath, DispatchRevDirPath, DispatchIssueDirPath, DispatchVolDirPath,  
                EproofNormal, ErrorDirPath, CommentEnableDir, DispatchUploadIn, DispatchSproofDirPath, SiteId, DispatchMenuItem, 
                GraphicsOnlineDirPath, GraphicsWebDirPath, MarkPDFDirPath, IssueManagement, WorkingFolderDirPath) values 
                ('" + scusttype + "','" + sCustSN + "','CG001','|" + TeamID + "|','" + Session["Job_ID"].ToString() + "','|Sample-Sample|FirstProof-FP|Revises-Rev|Finals-Fin|Issue-Iss|Online-Onl|'," +
            "'|WIP-WIP|FirstProof-FP|Revises-Rev|Finals-Fin|CE-CER|RE-RER|Issue-Iss|Issue Report-IssR|Online-Onl|','|Revises-Rev|Finals-Fin|FPP-FPP|FM/BM Details-FMBM|Online-Onl|Issue-Iss|'," +
            "'|All-All|FP-FP|Revises-Rev|Finals-Fin|Online-Onl|','1','4','4','" + strCustId + "'," +
            "'ML','F20-01','F20-02','F20-03','F40-01','F30-02','F50-02','F60-02','F60-05','F70-01','F70-03','F250-01','F250-02','F250-03','F250-04','F260-01'" +
            ",'F260-02','F270-01','F250-05','F250-06','L0002','|First Proof-FP|Revises-Rev|ePub-ePub|Finals-Fin|AM-AM|','F90-01','F90-02','F5000-04','1','F5000-08')";

                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    CustID = strCustId;
                    Session["FTPCustID"] = CustID;

                    strQuery = "Insert into JBM_CustomerMaster(CustID,CustName,CustCategory,RootID,CustSN,CustType,JBM_TeamID,RootIDCeninw,Cust_Disabled,JBM_Iss_UnAssigned) values('" + CustID + "','" + scustname + "','" + sCategory.Trim() + "','4','" + sCustSN + "','" + scusttype + "','" + TeamID + "','4','" + sCustomerStatus + "','"+Iss+"')";
                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                }
                else
                {
                    strQuery = "update JBM_CustomerMaster set CustName='" + scustname + "',CustCategory='" + sCategory.Trim() + "',JBM_Iss_UnAssigned='"+Iss+"' where CustID='" + sCustID + "'";
                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult UpdateCommCustomer(string sCountry, string sState, string sCity, string sAddress1, string sAddress2, string sEmail, string sPhone, string sFax, string sUrl,string sCustCode)
        {
            try
            {
                string CustID = "";
                if (Session["sStatus"] == null)
                    CustID = Session["FTPCustID"].ToString();
                else
                    CustID = sCustCode;
                string strQuery = @"Update JBM_CustomerMaster set CustAddress1='" + sAddress1 + "',CustAddress2='" + sAddress2 + "'," +
                    "CustCity='" + sCity + "',CustState='" + sState + "',CustCountry='" + sCountry + "',CustFax='" + sFax + "'," +
                    "CustPhone='" + sPhone + "',CustEmail='" + sEmail + "',CustWebAddress='" + sUrl + "' " +
                    " where CustID= '"+ CustID + "'";
                string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                                              

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult SendMail(string sFrom,string sTo,string sCC,  string sSubject, string sBody, string sBCC,string scusttype)
        {
            try
            {
                string strResult2;
                if (Session["sStatus"] == null)
                {
                    string strQuery = @"Insert into JBM_FTP_Stage_Details(CustType, JBM_AutoID, Stage, MailFrom, MailTo,
                MailCc, MailBcc, MailSubject, MailBody, MailErrorTo, MailErrorCc) values('" + scusttype + "','" + scusttype + "-Amail','Bookin','" + sFrom + "','" + sTo + "','" + sCC + "'," +
                    "'" + sBCC + "','108','108','MIS.Support@kwglobal.com','thiyagarajan.g@kwglobal.com')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    strQuery = @"Insert into JBM_MailInfo(EmpAutoID, JBM_AutoID, AutoArtID, MailEventID, MailStatus, 
                MailInitDate, Stage, MaxTry, First_try, MailFrom, MailTo, MailCc, MailBCc,
                MailSub, MailBody, IsBodyHtml, MailPriority, MailType, CustType) values('" + Session["EmpAutoId"] + "','','','E0001','1','" + DateTime.Now + "','','','','" + sFrom + "'," +
                   ",'" + sTo + "','" + sCC + "','" + sBCC + "','" + sSubject + "','" + sBody + "','1','2','Direct','" + scusttype + "')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    strQuery = @"Insert into JBM_FTP_Stage_Details(CustType, JBM_AutoID, Stage, MailFrom, MailTo,
                MailCc, MailBcc, MailSubject, MailBody, MailErrorTo, MailErrorCc) values('" + scusttype + "','" + scusttype + "-Amail','Bookin','" + sFrom + "','vijay.tejavat@kwglobal.com','<Rowena.Dsouza@kwglobal.com>; <Sivagnanamoorthy@kwglobal.com>; <Mayavan.Renganathan@kwglobal.com>; <Nirmal.Kumar@kwglobal.com>;  <thiyagarajan.g@kwglobal.com>; <aparna.MR@kwglobal.com>;<Sorimuthu.G@kwglobal.com>'," +
                   "'','109','109','MIS.Support@kwglobal.com','thiyagarajan.g@kwglobal.com')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    string strQueryFinal = "select MsgBody,Msgsubject from JBM_messageInfo where MsgID=109";
                    DataTable dt = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                    string mailSubject = "";
                    string mailBody = "";
                    if (dt.Rows.Count > 0)
                    {
                        mailSubject = dt.Rows[0]["Msgsubject"].ToString();
                        mailBody = dt.Rows[0]["MsgBody"].ToString().Replace("###EMPNAME###", Session["EmpName"].ToString()).Replace("###CUSTTYPE###", scusttype);
                    }
                    strQuery = @"Insert into JBM_MailInfo(EmpAutoID, JBM_AutoID, AutoArtID, MailEventID, MailStatus, 
                MailInitDate, Stage, MaxTry, First_try, MailFrom, MailTo, MailCc, MailBCc,
                MailSub, MailBody, IsBodyHtml, MailPriority, MailType, CustType) values('" + Session["EmpAutoId"] + "','','','E0001','1','" + DateTime.Now + "','','','','" + sFrom + "'," +
                   ",'vijay.tejavat@kwglobal.com','<Rowena.Dsouza@kwglobal.com>; <Sivagnanamoorthy@kwglobal.com>; <Mayavan.Renganathan@kwglobal.com>; <Nirmal.Kumar@kwglobal.com>;  <thiyagarajan.g@kwglobal.com>; <aparna.MR@kwglobal.com>;<Sorimuthu.G@kwglobal.com>',''," +
                   "'" + mailSubject + "','" + mailBody + "','1','2','Direct','" + scusttype + "')";
                     strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                }
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult FTPDetails()
        {
            Session["PrjTabHead"] = "Customer Information";
            Session["PageHead"] = "Customer Information";
            return View();
        }

        [SessionExpire]
        public ActionResult CustomerDetails()
        {
            try
            {
                //string nextStage = "";
                //nextStage = Proc_GetNextStage("C1", "S1131", "S1010");
                               
                List<SelectListItem> lstCustomer = new List<SelectListItem>();
                DataTable ds = new DataTable();
                string strQueryFinal = @"select *,(select EmpName from JBM_employeemaster where EmpAutoID= x.ApprovalBy) as ApprovalName,
(select EmpName from JBM_employeemaster where EmpAutoID = x.ApprovedBy) as ApprovedName
from(SELECT   CustID, CustName, CustCategory, CustAddress1, CustAddress2, CustCity, CustState, CustCountry, CustFax, CustPhone,
CustEmail, RootID, CustWebAddress, CustSN, CustType, JBM_TeamID, RootIDCeninw, Cust_Disabled, AccountMgr, ProjectMgr, SiteLeader,
 STSupportMgr, Status, ArtStageTypeID, WF_Code,
 (select EmpAutoID from JBM_employeemaster where EmpAutoID =
 case when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is not NULL and STSupportMgr is not NULL
 and Status = 'Approved' then STSupportMgr
 when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is not NULL and STSupportMgr is NULL
 and Status = 'Approved' then SiteLeader
 when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is NULL and STSupportMgr is NULL
 and Status = 'Approved' then ProjectMgr
 when AccountMgr is not NULL and ProjectMgr is NULL and SiteLeader is NULL and STSupportMgr is NULL
 and Status = 'Approved' then AccountMgr else Null End) as ApprovedBy, 
 (select EmpAutoID from JBM_employeemaster where EmpAutoID =
 case when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is not NULL and STSupportMgr is not NULL
 and Status = 'Pending' then STSupportMgr
 when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is not NULL and STSupportMgr is NULL
 and Status = 'Pending' then SiteLeader
 when AccountMgr is not NULL and ProjectMgr is not NULL and SiteLeader is NULL and STSupportMgr is NULL
 and Status = 'Pending' then ProjectMgr
 when AccountMgr is not NULL and ProjectMgr is NULL and SiteLeader is NULL and STSupportMgr is NULL
 and Status = 'Pending' then AccountMgr End) as ApprovalBy
 FROM JBM_CustomerMaster where Cust_Disabled is null or Cust_Disabled='0' ";
                if (Session["CustAccess"] != null)
                    strQueryFinal += " and CustType like '" + Session["CustAccess"].ToString().Trim() + "%'";
                strQueryFinal +=")x order by CustSN asc";

               ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                if (ds.Rows.Count > 0)
                {
                    ViewData.Model = ds.AsEnumerable();
                }

                Session["CustomerSN"] = "";
                //ViewBag.Customerlist = lstCustomer;

                //Load Approve List
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                if (Session["RoleID"].ToString().Trim() == "104")
                    dsPM = DBProc.GetResultasDataSet("select EmpAutoID,EmpName from JBM_employeemaster where roleid='103'", Session["sConnSiteDB"].ToString());
                else if (Session["RoleID"].ToString().Trim() == "103")
                    dsPM = DBProc.GetResultasDataSet("select EmpAutoID,EmpName from JBM_employeemaster where roleid='106'", Session["sConnSiteDB"].ToString());
                else if (Session["RoleID"].ToString().Trim() == "106")
                    dsPM = DBProc.GetResultasDataSet("select EmpAutoID,EmpName from JBM_employeemaster where roleid='100'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables.Count > 0)
                {
                    if (dsPM.Tables[0].Rows.Count > 0)
                    {
                        for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                        {
                            string strEmpAutoID = dsPM.Tables[0].Rows[intCount]["EmpAutoID"].ToString();
                            string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                            lstPM.Add(new SelectListItem
                            {
                                Text = strEmpName.ToString(),
                                Value = strEmpAutoID.ToString()
                            });
                        }
                    }
                }

                ViewBag.PMlist = lstPM;

                Session["PrjTabHead"] = "Customer Information";
                Session["PageHead"] = "Customer Information";
                return View();
            }
            catch (Exception ex)
            {
                return View();
            }
           
        }
        [SessionExpire]
        public ActionResult UpdateCustomerStatus(string sPending,string CID)
        {
            try
            {
                string strResult2;
                string strstart = "", strend = "", sstrStatus = ""; ;
                string strJM = "S1131";
                string strSL = "S1210";
                string EmpHeadName = "";
                string strStatus = "Pending";
                if (Session["RoleID"].ToString().Trim() == "104")
                {
                    strstart = Init_Barcode.gstrProjectManagement;
                    strend = strJM;
                    sstrStatus = Init_Barcode.gstrArtCompleted;
                    EmpHeadName = "AccountMgr";
                }
                else if (Session["RoleID"].ToString().Trim() == "103")
                {
                    strstart = strJM;
                    strend = strSL;
                    sstrStatus = Init_Barcode.gstrArtCompleted;
                    EmpHeadName = "ProjectMgr";
                }
                else if (Session["RoleID"].ToString().Trim() == "106")
                {
                    strstart = strSL;
                    strend = Init_Barcode.gstrTechSupport;
                    sstrStatus = Init_Barcode.gstrArtPending;
                    EmpHeadName = "SiteLeader";
                }
                else if (Session["RoleID"].ToString().Trim() == "100")
                {
                    strstart = strSL;
                    strend = Init_Barcode.gstrTechSupport;
                    sstrStatus = "S1211";
                    EmpHeadName = "STSupportMgr";
                    strStatus = "Approved";
                }

                string strArtStageTypeID = "";
               
                if (sPending.Trim() != "")
                {
                    strArtStageTypeID = strstart + "_" + sstrStatus + "_" + strend;
                }            


                string Squery = "";
                if(sPending!="")
                {
                    if (sPending == "1")
                    {
                        sPending = Session["EmpAutoId"].ToString();
                    }
                    Squery = "update JBM_CustomerMaster set Status='" + strStatus + "',ArtStageTypeID='" + strArtStageTypeID + "'," + EmpHeadName + "='" + sPending + "' where CustID='" + CID + "'";
                    string Strquery = DBProc.GetResultasString(Squery, Session["sConnSiteDB"].ToString());
                }
                DataTable dt = new DataTable();
                DataTable dtemp = new DataTable();
                string strcust = "select CustName,CustType,CustSN from JBM_CustomerMaster where CustID='" + CID + "'";
                dt = DBProc.GetResultasDataTbl(strcust, Session["sConnSiteDB"].ToString());

                strcust = "select EmpMailId,EmpName from JBM_employeemaster where EmpAutoID='" + sPending + "'";
                dtemp = DBProc.GetResultasDataTbl(strcust, Session["sConnSiteDB"].ToString());

                string strQueryFinal = "select MsgBody,Msgsubject from JBM_messageInfo where MsgID=108";
                DataTable dtmail = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                string mailSubject = "";
                string mailBody = "";
                if (dtmail.Rows.Count > 0)
                {

                    mailSubject = dtmail.Rows[0]["Msgsubject"].ToString().Replace("###CUSTNAME###", dt.Rows[0]["CustName"].ToString());
                    mailBody = dtmail.Rows[0]["MsgBody"].ToString().Replace("###EMPNAME###", dtemp.Rows[0]["EmpName"].ToString()).Replace("###DATE###", DateTime.Now.ToString()).Replace("###CUSTNAME###", dt.Rows[0]["CustName"].ToString()).Replace("###CUSTTYPE###", dt.Rows[0]["CustType"].ToString()).Replace("###CUSTSN###", dt.Rows[0]["CustSN"].ToString());
                }

                string strQuery = @"Insert into JBM_FTP_Stage_Details(CustType, JBM_AutoID, Stage, MailFrom, MailTo,
                MailCc, MailBcc, MailSubject, MailBody, MailErrorTo, MailErrorCc) values('" + dt.Rows[0]["CustType"].ToString() + "','" + dt.Rows[0]["CustType"].ToString() + "-Amail','Bookin','" + Session["EmpMailId"].ToString() + "','" + dtemp.Rows[0]["EmpMailId"].ToString() + "',''," +
               "'','108','108','MIS.Support@kwglobal.com','thiyagarajan.g@kwglobal.com')";
                 strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                strQuery = @"Insert into JBM_MailInfo(EmpAutoID, JBM_AutoID, AutoArtID, MailEventID, MailStatus, 
                MailInitDate, Stage, MaxTry, First_try, MailFrom, MailTo, MailCc, MailBCc,
                MailSub, MailBody, IsBodyHtml, MailPriority, MailType, CustType) values('" + Session["EmpAutoId"] + "','','','E0001','1','" + DateTime.Now + "','','','','" + Session["EmpMailId"].ToString() + "'," +
               ",'" + dtemp.Rows[0]["EmpMailId"].ToString() + "','','','" + mailSubject + "','" + mailBody + "','1','2','Direct','" + dt.Rows[0]["CustType"].ToString() + "')";
                 strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                   
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult ViewCustomerDetails(string sCustID)
        {
            try
            {
                if (sCustID != "")
                {
                    Session["sesCustID"] = sCustID;
                    Session["sesStatus"] = "View";
                    Session["PageSubHead"] = "Customer Details";
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult EditCustomerDetails(string sCustID)
        {
            try
            {                
                if (sCustID != "")
                {
                    DataTable dt = DBProc.GetResultasDataTbl("Select CustType from JBM_CustomerMaster where CustID='" + sCustID + "'", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        Session["sCustAcc"] = dt.Rows[0]["CustType"].ToString();
                    }

                    Session["sesCustID"] = sCustID;
                    Session["sesStatus"] = "Edit";
                    Session["PageSubHead"] = "Update Customer";
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult UpdateCustomerDetails(string sCustID,string sCountry, string sState, string sCity, string sCategory, string sKglteam, string sCenproPath, string sCeninwPath, string scustname, string scusttype, string sAddress1, string sAddress2, string sCustSN,  string sEmail, string sPhone, string sFax, string sUrl, string sProfileName, string sFtphost, string sFtppath, string sUserid, string sPwd, string sCustomerStatus, string sCustAccess, string sStage, string sDestinationPath, string sCutorCopy)
        {
            try
            {
                if (sCustID != "")
                {
                    string strQuery = "update JBM_CustomerMaster set CustName='" + scustname + "',CustCategory='" + sCategory.Trim() + "',CustAddress1='" + sAddress1 + "',CustAddress2='" + sAddress2 + "',CustCity='" + sCity + "',CustState='" + sState + "',CustCountry='" + sCountry.Trim() + "',CustFax='" + sFax + "',CustPhone='" + sPhone + "',CustEmail='" + sEmail + "',RootID='" + sCenproPath + "',CustWebAddress='" + sUrl + "',CustSN='" + sCustSN + "',CustType='" + scusttype + "',JBM_TeamID='" + sKglteam + "',RootIDCeninw='" + sCeninwPath + "',Cust_Disabled='" + sCustomerStatus + "' where CustID='" + sCustID + "'";
                    string strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                    strQuery = "update JBM_Ftp_Details set ftpProfileName='" + sProfileName + "',ftpHost='" + sFtphost + "',ftpUID='" + sUserid + "',ftpPWD='" + sPwd + "',ftpPath='" + sFtppath + "',CustAccess='" + sCustAccess + "',Stage='" + sStage + "',DestinationPath='" + sDestinationPath + "',CutorCopy='" + sCutorCopy + "' where CustID='" + sCustID + "'";
                    strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                    
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
           
        }
        public string CreateBtn(string uniqueID)
        {
            string formControl = string.Empty;
            try 
            {
                if (uniqueID != "")
                {
                    formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\"  class='btn btn-light' name='edit' value='Edit'><span class='fas fa-edit fa-1x text-green'></span></button><button type='button' id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\"  class='btn btn-light' name='delete' value='Delete'><span class='fas fa-trash fa-1x text-red'></span></button>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        private string Proc_GetNextStage(string strWFCode, string strCurrentActivity, string strCurrentStatus)
        {
            string[] strArray;
            string strWFCodeXML = string.Empty;
            strArray = GetAllKeysInIniFileSection(strWFCode, @"C:\inetpub\WorkingFolder\Init\WorkflowCode.xml");
            for (int i = 0; i < strArray.Length; i++)
            {
                strWFCodeXML += string.Join("\r\n", strArray[i]);
            }
            string strNxtStage;
            strNxtStage = Regex.Match(strWFCodeXML, "<ArtNextStage Val=\"" + strCurrentActivity + "_" + strCurrentStatus + "_" + "(.*?)\" />").Groups[1].Value;
            //Proc_GetNextStage = strNxtStage;
            return strNxtStage;
        }
        //bool Proc_Get_NextStage_Info(string strAutoArtID, string strWorkflowCode, string strDeptAct, string strStatus, ref string strSelAct, string strSelProcess, ref string strNxtStage, ref string strArtStage, ref ArtDet A, ref string ceFrelancer, string strRTPProcess, string strM2c, string strCurrentProcess)
        //{
        //    try
        //    {
        //        string strDeptCode = Session["DeptCode"].ToString();
        //        string strUserId = Session["EmpAutoId"].ToString();
        //        string strUserLogin = Session["EmpLogin"].ToString();
        //        string strCustAccess = Session["sCustAcc"].ToString();
        //        string xmlString="";

        //        if (Request.Url.ToString().ToLower().Contains("localhost"))
        //        {
        //            // System.IO.File.Open(@"C:\inetpub\WorkingFolder\Init\BaseWorkflow.xml", FileMode.Open, FileAccess.Read, FileShare.ReadWrite); 
        //            XmlDocument doc = new XmlDocument();
        //            doc.Load(@"C:\inetpub\WorkingFolder\Init\BaseWorkflow.xml");
        //            foreach (XmlNode node in doc.DocumentElement.ChildNodes)
        //            {
        //                string text = node.InnerText; //or loop through its children as well
        //            }
        //            xmlString = System.IO.File.ReadAllText(@"C:\inetpub\WorkingFolder\Init\BaseWorkflow.xml");
        //        }
        //        else
        //        {
        //            System.IO.File.Open(@"C:\inetpub\WorkingFolder\Init\BaseWorkflow.xml", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        //        }

        //        string gBarActXML = xmlString;// System.IO.File.InputString(1, FileSystem.LOF(1));
        //        //FileSystem.FileClose(1);

        //        string strWFCode = string.Empty;
        //        string[] strArray;
        //        if (Request.Url.ToString().ToLower().Contains("localhost"))
        //            strArray = GetAllKeysInIniFileSection(strWorkflowCode, @"C:\inetpub\WorkingFolder\Init\WorkflowCode.xml");
        //        else
        //            strArray = GetAllKeysInIniFileSection(strWorkflowCode, @"C:\inetpub\WorkingFolder\Init\WorkflowCode.xml");
        //        for (int i = 0; i < strArray.Length; i++)
        //        {
        //            strWFCode = string.Join("\r\n", strArray[i]);
        //        }

        //        Regex RegEx1;
        //        strNxtStage = Regex.Match(strWFCode, "<ArtNextStage Val=\"" + strDeptAct + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value;

        //        if (strNxtStage == "")
        //        {
        //            gBarActXML = gBarActXML.Replace("</Stages>", strWFCode + "</Stages>");
        //            string strPattern = "<Workflow>.*?<ArtStageDes>(.*?)</ArtStageDes>.*?<Stages>(.*?)</Stages>.*?</Workflow>";
        //            RegEx1 = new Regex(strPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
        //            Match m = RegEx1.Match(gBarActXML);
        //            if (m.Success)
        //            {
        //                string strTemp = m.Groups[2].ToString();
        //                if (strDeptAct == Init_Barcode.gstrIM & strSelAct == Init_Barcode.gstrPass2Login)
        //                {
        //                    // process change for IM passing the article direct to comp
        //                    strDeptAct = Init_Barcode.gstrLogin;
        //                    strSelAct = Init_Barcode.gstrArtCompleted;
        //                    strStatus = Init_Barcode.gstrArtCompleted;
        //                }
        //                strNxtStage = Regex.Match(strTemp, "<ArtNextStage Val=\"" + strDeptAct + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value;
        //            }
        //        }

        //        strArtStage = strDeptAct + "_" + strStatus + "_";

        //        if (strStatus == Init_Barcode.gstrArtRejected)
        //        {
        //            DataTable dt = DBProc.GetResultasDataTbl("Select ArtStagePrev from BK_ProdStatus where empAutoID='" + strUserId + "' and AutoArtID='" + strAutoArtID + "' and ArtIndex='1'", Session["sConnSiteDB"].ToString());// RecordManager.GetRecord_Multiple_All(, "TempTable");
        //            if (!(dt.Rows[0]["ArtStagePrev"].ToString() == null) & dt.Rows[0]["ArtStagePrev"].ToString() != "")
        //                strNxtStage = dt.Rows[0]["ArtStagePrev"].ToString();
        //        }
        //        if (strDeptCode == Init_Department.gDcProjectManagement & strStatus == Init_Barcode.gstrArtCompleted & A.SD.blnCCHighSpeedSuccess & A.AD.Stage.IndexOf("RP") == -1)
        //        {
        //            if ((Regex.Match(strWFCode, "<ArtNextStage Val=\"" + Init_Barcode.gstrLogin + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value) != "")
        //                strNxtStage = Regex.Match(strWFCode, "<ArtNextStage Val=\"" + Init_Barcode.gstrLogin + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value;
        //        }
        //        if (strStatus == Init_Barcode.gstrSE2xStyle & strCurrentProcess == Init_Barcode.gstrSE2xStyle & strDeptCode == Init_Department.gDcRE)
        //        {
        //            strNxtStage = Init_Barcode.gstrSE2xStyle;
        //            strArtStage = strDeptAct + "_" + Init_Barcode.gstrArtCompleted + "_" + Init_Barcode.gstrSE2xStyle + "_" + strUserLogin + "_";
        //        }
        //        if (strNxtStage == Init_Barcode.gstrConversion & A.SD.blnCCHighSpeedSuccess & strDeptCode == Init_Department.gDcProjectManagement & strStatus == Init_Barcode.gstrArtCompleted)
        //        {
        //            if ((Regex.Match(strWFCode, "<ArtNextStage Val=\"" + Init_Barcode.gstrConversion + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value) != "")
        //                strNxtStage = Regex.Match(strWFCode, "<ArtNextStage Val=\"" + Init_Barcode.gstrConversion + "_" + strSelAct + "_" + "(.*?)\" />").Groups[1].Value;
        //        }
        //        if (strDeptCode == Init_Department.gDcPag & strM2c == "TeX-CER" & strStatus == Init_Barcode.gstrDispatch)
        //        {
        //            strArtStage = strDeptAct + "_" + Init_Barcode.gstrArtCompleted + "_" + Init_Barcode.gstrLaTeX2HTML + "_" + strUserLogin + "_";
        //            strNxtStage = Init_Barcode.gstrLaTeX2HTML;
        //        }
        //        if (strNxtStage != "")
        //        {
        //            if (strDeptCode == Init_Department.gDcCE & strStatus == Init_Barcode.gstrArtCompleted & strSelProcess == "Art Corr.")
        //                strArtStage += strDeptAct + "_" + strUserLogin + "_";
        //            else if (strCustAccess == "OP" & ceFrelancer == "2" & strStatus == Init_Barcode.gstrArtPass2IM)
        //            {
        //                strNxtStage = Init_Barcode.gstrPass2IMforFreelancer;
        //                strArtStage += strNxtStage + "_" + strUserLogin + "_";
        //            }
        //            else if (strCustAccess == "OP" & A.AD.DeptCode == "20" & strNxtStage == "S1002")
        //            {
        //                strNxtStage = Init_Barcode.gstrCleanUp;
        //                A.AD.SelProcess = "CleanUp";
        //                // strArtStage &= strNxtStage & "_" & strUserLogin & "_"
        //                strSelAct = Init_Barcode.gstrCleanUp;
        //                strArtStage = Init_Barcode.gstrCeRapid + "_" + Init_Barcode.gstrCleanUp + "_" + Init_Barcode.gstrCleanUp + "_" + strUserLogin + "_";
        //            }
        //            else if (strDeptCode == Init_Department.gDcPag & strStatus == Init_Barcode.gstrArtCompleted & strRTPProcess == Init_Barcode.gstrRTP)
        //            {
        //                strNxtStage = Init_Barcode.gstrIM;
        //                strArtStage = strDeptAct + "_" + Init_Barcode.gstrArtCompleted + "_" + Init_Barcode.gstrIM + "_" + strUserLogin + "_";
        //            }
        //            else if (strCustAccess == "OP" & strDeptCode == Init_Department.gDcCE & strStatus == Init_Barcode.gstrArtCompleted & strNxtStage == Init_Barcode.gstrCopyFilesforReview)
        //                strNxtStage = Init_Barcode.gstrXML;
        //            else if (strStatus == Init_Barcode.gstrArtCompleted & strCurrentProcess == Init_Barcode.gstrMechanicalEditing & strDeptCode == Init_Department.gDcRE & strCustAccess == "CS")
        //            {
        //                strNxtStage = Init_Barcode.gstrPE;
        //                strArtStage += strNxtStage + "_" + strUserLogin + "_";
        //            }
        //            else if (strStatus == Init_Barcode.gstrSE2xStyle & strCurrentProcess == Init_Barcode.gstrSE2xStyle & strDeptCode == Init_Department.gDcRE)
        //            {
        //                strNxtStage = Init_Barcode.gstrSE2xStyle;
        //                // strArtStage = Init_Barcode.gstrSE2xStyle & "_" & Init_Barcode.gstrArtCompleted & "_" & strNxtStage & "_" & strUserLogin & "_"
        //                strArtStage = strDeptAct + "_" + Init_Barcode.gstrArtCompleted + "_" + Init_Barcode.gstrSE2xStyle + "_" + strUserLogin + "_";
        //            }
        //            else
        //                strArtStage += strNxtStage + "_" + strUserLogin + "_";
        //        }
        //        else
        //        {
        //            //Proc_Err_Label_Update_Dropdown(rCboBox, "Missing Next Stage information in the workflow. Please check with Software Dept.");
        //            return false;
        //        }
        //        A.AD.NextStage = strNxtStage;
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        private static string[] GetAllKeysInIniFileSection(string strSectionName, string strIniFileName)
        {
            // Allocate in unmanaged memory a buffer of suitable size.
            // I have specified here the max size of 32767 as documentated
            // in MSDN.
            IntPtr pBuffer = Marshal.AllocHGlobal(32767);
            // Start with an array of 1 string only.
            // Will embellish as we go along.
            string[] strArray = new string[0];
            UInt32 uiNumCharCopied = 0;

            uiNumCharCopied = GetPrivateProfileSection(strSectionName, pBuffer, 32767, strIniFileName);

            // iStartAddress will point to the first character of the buffer,
            int iStartAddress = pBuffer.ToInt32();
            // iEndAddress will point to the last null char in the buffer.
            int iEndAddress = iStartAddress + (int)uiNumCharCopied;

            // Navigate through pBuffer.
            while (iStartAddress < iEndAddress)
            {
                // Determine the current size of the array.
                int iArrayCurrentSize = strArray.Length;
                // Increment the size of the string array by 1.
                Array.Resize(ref strArray, iArrayCurrentSize + 1);
                // Get the current string which starts at "iStartAddress".
                string strCurrent = Marshal.PtrToStringAnsi(new IntPtr(iStartAddress));
                // Insert "strCurrent" into the string array.
                strArray[iArrayCurrentSize] = strCurrent;
                // Make "iStartAddress" point to the next string.
                iStartAddress += (strCurrent.Length + 1);
            }

            Marshal.FreeHGlobal(pBuffer);
            pBuffer = IntPtr.Zero;

            return strArray;
        }

    }
}

