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
//using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace SmartTrack.Controllers
{
   // [SessionExpire]
    public class ProjectTrackController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();

        // GET: ProjectTrack
        public ActionResult Index(string CustAcc, string EmpID, string SiteID)
        {

            DataTable dtnew = new DataTable();
            if (CustAcc != null)
            {
                string strUrl = Request.Url.AbsoluteUri.ToString();

                Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";
                Session["strHomeURL"] = strUrl;
                Session["EmpIdLogin"] = EmpID;
                Session["sCustAcc"] = CustAcc;
                Session["sSiteID"] = SiteID;
                Session["CustomerSN"] = "";
                clsCollec.getSiteDBConnection(SiteID, CustAcc);
                if (Session["sConnSiteDB"].ToString() == "")
                {
                    Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                }

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select EmpLogin,EmpAutoId,EmpName,EmpMailId,DeptCode,RoleId, DeptAccess,CustAccess,TeamID,(Select b.DeptName from JBM_DepartmentMaster b  where e.DeptCode = b.DeptCode) as DeptName from JBM_EmployeeMaster e Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
                    Session["EmpLogin"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
                    Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();
                    Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                    Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                    Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();

                    GlobalVariables.strEmpName = ds.Tables[0].Rows[0]["EmpName"].ToString();
                }
            }
            ViewBag.PrjTabHead = "Project Tracking";
            ViewBag.PageHead = "Project Tracking";
            return View();
        }
        [SessionExpire]
        public ActionResult ProjectTracking(string CustAcc, string EmpID, string SiteID,string sCustSN,string CustId,string uniqueID,string uniqueName,string strdashboard,string sStatus,string sHealth, string strStatus)
        {
            try
            {
                DataTable dtnew = new DataTable();
                if (CustAcc != null)
                {
                    string strUrl = Request.Url.AbsoluteUri.ToString();

                    Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";
                    Session["strHomeURL"] = strUrl;
                    Session["EmpIdLogin"] = EmpID;
                    Session["sCustAcc"] = CustAcc;
                    Session["sSiteID"] = SiteID;
                    Session["CustomerSN"] = "";

                    if (Session["PageLength"] == null)
                    {
                        Session["PageLength"] = 10;
                    }
                    if (Session["sPMProject"] == null)
                    {
                        Session["sPMProject"] = "MyProjects";
                    }
                    
                    clsCollec.getSiteDBConnection(SiteID, CustAcc);
                    if (Session["sConnSiteDB"].ToString() == "")
                    {
                        Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                    }

                    DataSet dss = new DataSet();
                    dss = DBProc.GetResultasDataSet("Select EmpLogin,EmpAutoId,EmpName,EmpMailId,DeptCode,RoleId, DeptAccess,CustAccess,TeamID,(Select b.DeptName from JBM_DepartmentMaster b  where e.DeptCode = b.DeptCode) as DeptName from JBM_EmployeeMaster e Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
                    if (dss.Tables[0].Rows.Count > 0)
                    {
                        Session["EmpAutoId"] = dss.Tables[0].Rows[0]["EmpAutoId"].ToString();
                        Session["EmpLogin"] = dss.Tables[0].Rows[0]["EmpLogin"].ToString();
                        Session["EmpName"] = dss.Tables[0].Rows[0]["EmpName"].ToString();
                        Session["DeptName"] = dss.Tables[0].Rows[0]["DeptName"].ToString();
                        Session["DeptCode"] = dss.Tables[0].Rows[0]["DeptCode"].ToString();
                        Session["RoleID"] = dss.Tables[0].Rows[0]["RoleID"].ToString();

                        GlobalVariables.strEmpName = dss.Tables[0].Rows[0]["EmpName"].ToString();
                    }
                }

                Session["CustID"] = "";
                if (Session["RoleID"].ToString() == "102" || Session["CustomerSN"].ToString() !="" || strdashboard == "Projects" || strdashboard == "DashBoard")
                {
                    DataSet ds1 = new DataSet();
                    string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and jc.CustSN='" + sCustSN + "' order by CustSN asc ";
                    ds1 = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        ViewBag.vCustID = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        ViewBag.vCustSN = ds1.Tables[0].Rows[0]["CustSN"].ToString();
                        ViewBag.vCustName = ds1.Tables[0].Rows[0]["CustName"].ToString();
                        Session["CustID"] = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        Session["CustomerName"] = ds1.Tables[0].Rows[0]["CustName"].ToString();
                    }

                    Session["CustomerSN"] = sCustSN;
                }
                Session["sesSite"] = "";
                
                if (strdashboard == "DashBoard")
                {
                    ViewBag.strStatus = strStatus;
                    ViewBag.sHealth = sHealth;
                    ViewBag.strdashboard = strdashboard;
                    ViewBag.uniqueName = uniqueName;
                    ViewBag.sStatus = sStatus;
                    if (uniqueName == "Service")
                        ViewBag.ServiceID = uniqueID;
                    else if (uniqueName == "PM")
                        ViewBag.PMID = uniqueID;
                    else if (uniqueName == "CPM")
                        ViewBag.CPMID = uniqueID;
                    else if (uniqueName == "Division")
                        ViewBag.DivID = uniqueID;
                    else if (uniqueName == "Site")
                    {
                        ViewBag.SiteID = uniqueID;
                        Session["sesSite"] = uniqueID;
                    }
                }
                else if (strdashboard == "Projects")
                {
                    ViewBag.strdashboard = strdashboard;
                    ViewBag.sCstSN = sCustSN;
                }
                else
                {
                    ViewBag.strdashboard = "";
                    ViewBag.uniqueName = "";
                    ViewBag.sStatus = "";
                    ViewBag.sCstSN = "";
                }
                //Load Project list Items
                List<SelectListItem> lstProject = new List<SelectListItem>();
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select DISTINCT jc.CustSN from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        string strCustSN = ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                        lstProject.Add(new SelectListItem
                        {
                            Text = strCustSN.ToString(),
                            Value = strCustSN.ToString()
                        });
                    }

                }
                
                ViewBag.Projectlist = lstProject;

                //Load Project Manager 
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                dsPM = DBProc.GetResultasDataSet("select EmpName,emplogin from JBM_Employeemaster where usergroup = (select emplogin from JBM_Employeemaster where emplogin = '" + Session["EmpLogin"].ToString().Trim() + "' and roleid = '103') and roleid = '104'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                        string stremplogin = dsPM.Tables[0].Rows[intCount]["emplogin"].ToString();

                        lstPM.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = stremplogin.ToString()
                        });
                    }

                }

                ViewBag.PMlist = lstPM;

                //Load Customer Project Manager 
                List<SelectListItem> lstCPM = new List<SelectListItem>();
                DataSet dsCPM = new DataSet();
                dsCPM = DBProc.GetResultasDataSet("SELECT ProjectManagerUS FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '"+ Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.ProjectManagerUS", Session["sConnSiteDB"].ToString());
                if (dsCPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strCustSN = dsCPM.Tables[0].Rows[intCount]["ProjectManagerUS"].ToString();
                        lstCPM.Add(new SelectListItem
                        {
                            Text = strCustSN.ToString(),
                            Value = strCustSN.ToString()
                        });
                    }

                }

                ViewBag.CPMlist = lstCPM;

                //Load Project Manager list Items
                List<SelectListItem> itemsProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpAutoID,EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='104' or roleid='103' order by empname asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.PrjMgrList = itemsProjectManager;

                ViewBag.PrjTabHead = "Projects";
                ViewBag.PageHead = "Project Tracking - Project List";
            }
            catch (Exception)
            {
                throw;
            }
            return View();
        }
        [SessionExpire]
        public ActionResult Reports(string CustAcc, string EmpID, string SiteID, string sCustSN, string CustId, string uniqueID, string uniqueName, string strdashboard, string sStatus, string sHealth, string strStatus)
        {
            try
            {
                DataTable dtnew = new DataTable();
                if (CustAcc != null)
                {
                    string strUrl = Request.Url.AbsoluteUri.ToString();

                    Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";
                    Session["strHomeURL"] = strUrl;
                    Session["EmpIdLogin"] = EmpID;
                    Session["sCustAcc"] = CustAcc;
                    Session["sSiteID"] = SiteID;
                    Session["CustomerSN"] = "";

                    if (Session["PageLength"] == null)
                    {
                        Session["PageLength"] = 10;
                    }

                    clsCollec.getSiteDBConnection(SiteID, CustAcc);
                    if (Session["sConnSiteDB"].ToString() == "")
                    {
                        Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                    }

                    DataSet dss = new DataSet();
                    dss = DBProc.GetResultasDataSet("Select EmpLogin,EmpAutoId,EmpName,EmpMailId,DeptCode,RoleId, DeptAccess,CustAccess,TeamID,(Select b.DeptName from JBM_DepartmentMaster b  where e.DeptCode = b.DeptCode) as DeptName from JBM_EmployeeMaster e Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
                    if (dss.Tables[0].Rows.Count > 0)
                    {
                        Session["EmpAutoId"] = dss.Tables[0].Rows[0]["EmpAutoId"].ToString();
                        Session["EmpLogin"] = dss.Tables[0].Rows[0]["EmpLogin"].ToString();
                        Session["EmpName"] = dss.Tables[0].Rows[0]["EmpName"].ToString();
                        Session["DeptName"] = dss.Tables[0].Rows[0]["DeptName"].ToString();
                        Session["DeptCode"] = dss.Tables[0].Rows[0]["DeptCode"].ToString();
                        Session["RoleID"] = dss.Tables[0].Rows[0]["RoleID"].ToString();

                        GlobalVariables.strEmpName = dss.Tables[0].Rows[0]["EmpName"].ToString();
                    }
                }

                Session["CustID"] = "";
                if (Session["RoleID"].ToString() == "102" || Session["CustomerSN"].ToString() != "" || strdashboard == "Projects" || strdashboard == "DashBoard")
                {
                    DataSet ds1 = new DataSet();
                    string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and jc.CustSN='" + sCustSN + "' order by CustSN asc ";
                    ds1 = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        ViewBag.vCustID = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        ViewBag.vCustSN = ds1.Tables[0].Rows[0]["CustSN"].ToString();
                        ViewBag.vCustName = ds1.Tables[0].Rows[0]["CustName"].ToString();
                        Session["CustID"] = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        Session["CustomerName"] = ds1.Tables[0].Rows[0]["CustName"].ToString();
                    }

                    Session["CustomerSN"] = sCustSN;
                }
                Session["sesSite"] = "";

                if (strdashboard == "DashBoard")
                {
                    ViewBag.strStatus = strStatus;
                    ViewBag.sHealth = sHealth;
                    ViewBag.strdashboard = strdashboard;
                    ViewBag.uniqueName = uniqueName;
                    ViewBag.sStatus = sStatus;
                    if (uniqueName == "Service")
                        ViewBag.ServiceID = uniqueID;
                    else if (uniqueName == "PM")
                        ViewBag.PMID = uniqueID;
                    else if (uniqueName == "CPM")
                        ViewBag.CPMID = uniqueID;
                    else if (uniqueName == "Division")
                        ViewBag.DivID = uniqueID;
                    else if (uniqueName == "Site")
                    {
                        ViewBag.SiteID = uniqueID;
                        Session["sesSite"] = uniqueID;
                    }
                }
                else if (strdashboard == "Projects")
                {
                    ViewBag.strdashboard = strdashboard;
                    ViewBag.sCstSN = sCustSN;
                }
                else
                {
                    ViewBag.strdashboard = "";
                    ViewBag.uniqueName = "";
                    ViewBag.sStatus = "";
                    ViewBag.sCstSN = "";
                }
                //Load Project list Items
                List<SelectListItem> lstProject = new List<SelectListItem>();
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select DISTINCT jc.CustSN from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        string strCustSN = ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                        lstProject.Add(new SelectListItem
                        {
                            Text = strCustSN.ToString(),
                            Value = strCustSN.ToString()
                        });
                    }

                }

                ViewBag.Projectlist = lstProject;

                //Load Project Manager 
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                dsPM = DBProc.GetResultasDataSet("select EmpName,emplogin from JBM_Employeemaster where usergroup = (select emplogin from JBM_Employeemaster where emplogin = '" + Session["EmpLogin"].ToString().Trim() + "' and roleid = '103') and roleid = '104'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                        string stremplogin = dsPM.Tables[0].Rows[intCount]["emplogin"].ToString();

                        lstPM.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = stremplogin.ToString()
                        });
                    }

                }

                ViewBag.PMlist = lstPM;

                //Load Customer Project Manager 
                List<SelectListItem> lstCPM = new List<SelectListItem>();
                DataSet dsCPM = new DataSet();
                dsCPM = DBProc.GetResultasDataSet("SELECT ProjectManagerUS FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.ProjectManagerUS", Session["sConnSiteDB"].ToString());
                if (dsCPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strCustSN = dsCPM.Tables[0].Rows[intCount]["ProjectManagerUS"].ToString();
                        lstCPM.Add(new SelectListItem
                        {
                            Text = strCustSN.ToString(),
                            Value = strCustSN.ToString()
                        });
                    }

                }

                ViewBag.CPMlist = lstCPM;
                //Load Project Manager list Items
                List<SelectListItem> itemsProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpAutoID,EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='104' or roleid='103' order by empname asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.PrjMgrList = itemsProjectManager;
                ViewBag.PrjTabHead = "Projects";
                ViewBag.PageHead = "Project Tracking - Export Reports";
            }
            catch (Exception)
            {
                throw;
            }
            return View();
        }

        [SessionExpire]
        public ActionResult GetPMList(string sCustSN, string sCPM, string sPMProject, string sServices, string sStatus, string sHealth, string sDivision, string sSite, string strHealth, string strrStatus)
        {
            try
            {
                string strRoleID = Session["RoleID"].ToString();
                if (Session["CustomerSN"].ToString() != "")
                {
                    sCustSN = Session["CustomerSN"].ToString();
                }

                string strQueryFinal = "SELECT distinct (select empname from JBM_employeemaster where emplogin=pm.ProjectCoordInd) as PM,pm.ProjectCoordInd as PMVal FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and pm.current_Status is not null and(select empname from JBM_employeemaster where emplogin = pm.ProjectCoordInd) is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";

                if (sCustSN != "AllCustomers")
                {
                    strQueryFinal += " and jc.CustSN in ('" + sCustSN.Trim() + "')";
                }

                if (sServices != "AllServices")
                {
                    strQueryFinal += " and ji.BM_FullService = '" + sServices.Trim() + "'";
                }
               
                if (sStatus != "AllStatus")
                {
                    if (sStatus != "DateRange")
                    {
                        if (strrStatus != "" && strrStatus != null)
                        {
                            strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                        }
                        else
                            strQueryFinal += " and pm.Current_Status = '" + sStatus.Trim() + "'";
                    }
                    else
                    {
                        if (strrStatus != "" && strrStatus != null)
                        {
                            strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                        }
                        strQueryFinal += " and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + Session["sStartDate"].ToString() + "', 101)  and CONVERT(DATETIME, '" + Session["sEndDate"].ToString() + "',101)";
                    }
                }
                else
                {
                    if (strrStatus != "" && strrStatus != null)
                    {
                        strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                    }
                }

                if (sHealth != "AllHealth")
                {
                    strQueryFinal += " and pm.Current_Health = '" + sHealth.Trim() + "'";
                }

                if (sCPM != "AllCPMs")
                {
                    strQueryFinal += " and pm.ProjectManagerUS = '" + sCPM.Trim() + "'";
                }               
                if (sDivision != "AllDivisions")
                {
                    if (sDivision == "N/A")
                        sDivision = "0";
                    else if (sDivision == "HSS")
                        sDivision = "1";
                    else if (sDivision == "STEM")
                        sDivision = "2";
                    strQueryFinal += " and ji.JBM_Location = '" + sDivision.Trim() + "'";
                }
                if (sSite != "")
                {
                    strQueryFinal += " and pm.Cenveo_Facility = '" + sSite.Trim() + "'";
                }
                if (strHealth != "" && strHealth != null)
                {
                    strQueryFinal += " and pm.Current_Health = '" + strHealth.Trim() + "'";
                }
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                List<SelectListItem> PMlst = new List<SelectListItem>();
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        string strPM = ds.Tables[0].Rows[intCount]["PM"].ToString(); 
                             string strPMVal = ds.Tables[0].Rows[intCount]["PMVal"].ToString();
                        PMlst.Add(new SelectListItem
                        {
                            Text = strPM.ToString(),
                            Value = strPMVal.ToString()
                        });
                    }

                }
                return Json(PMlst, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult SetSessionPagelength(string sPageLength,string sList)
        {
            if(sPageLength!="")
            {
                if(sList=="ProjectTracking")
                    Session["PageLength"] = sPageLength;
                else if (sList == "Schedule")
                    Session["SchedulePageLength"] = sPageLength;
                else if (sList == "Remarks")
                    Session["RemarksPageLength"] = sPageLength;                
                
            }
            return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
        }
        [SessionExpire]
        public ActionResult GetProjectList(string sCustSN, string sCPM, string sPMProject, string sServices, string sStatus, string sHealth, string sDivision, string sSite, string strHealth, string strrStatus)
        {
            try
            {
                Session["sPMProject"] = sPMProject;
                ViewBag.strdashboard = "";
                ViewBag.ServiceID = "";
                ViewBag.uniqueName = "";
                ViewBag.PMID = "";
                ViewBag.CPMID = "";
                ViewBag.SiteID = "";
                ViewBag.DivID = "";
                ViewBag.sStatus = "";
                ViewBag.sCstSN = "";
                ViewBag.sHealth = "";
                ViewBag.strStatus = "";

                string strRoleID = Session["RoleID"].ToString();
                if (Session["CustomerSN"].ToString() != "")
                {
                    sCustSN = Session["CustomerSN"].ToString();
                }

                //Session assign for to filter schedule and remarks tab
                Session["sesCustSN"] = sCustSN;
                Session["sesCPM"] = sCPM;
                Session["sesPMProject"] = sPMProject;
                Session["sesServices"] = sServices;

                if (strHealth != "" && strHealth != null)
                    Session["sesHealth"] = strHealth;
                else
                    Session["sesHealth"] = sHealth;
                if (strrStatus != "" && strrStatus != null)
                    Session["sesStatus"] = strrStatus;
                else
                    Session["sesStatus"] = sStatus;

                string strQueryFinal = "SELECT ji.JBM_AutoID,CASE WHEN ji.JBM_Location='0' THEN 'N/A' WHEN ji.JBM_Location='1' THEN 'HSS' WHEN ji.JBM_Location='2' THEN 'STEM' ELSE ji.JBM_Location END  As Division, pm.Current_Health, pm.Current_Status, jc.CustName, jc.CustSN, FORMAT(CAST(ji.JBM_PrinterDate as date), 'dd-MMM-yy') as JBM_PrinterDate, ji.JBM_ID, ji.Title, ji.BM_FullService, pm.ProjectManagerUS as CPM, (select empname from JBM_employeemaster where emplogin=pm.ProjectCoordInd) as PM,(case when (select sum (case when [ActualPages] is null then 0 else [ActualPages] end) as [ActualPages] From  BK_ChapterInfo WHERE JBM_AutoID=ji.JBM_AutoID) = 0 then (select BM_DesiredPgCount from JBM_Info where JBM_AutoID=ji.JBM_AutoID) else (select sum(case when[ActualPages] is null then 0 else [ActualPages] end) as [ActualPages] From BK_ChapterInfo WHERE JBM_AutoID = ji.JBM_AutoID) end) as [PgCount],ji.JBM_Intrnl,(ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)% 3 AS Col, (ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)/ 3 AS Row FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and   pm.current_Status is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                //string strQueryFinal = "SELECT ji.JBM_AutoID,CASE WHEN ji.JBM_Location='0' THEN 'N/A' WHEN ji.JBM_Location='1' THEN 'HSS' WHEN ji.JBM_Location='2' THEN 'STEM' ELSE ji.JBM_Location END  As Division, pm.Current_Health, pm.Current_Status, jc.CustName, jc.CustSN, FORMAT(CAST(ji.JBM_PrinterDate as date), 'dd-MMM-yy') as JBM_PrinterDate, ji.JBM_ID, ji.Title, ji.BM_FullService, pm.ProjectManagerUS as CPM, (select empname from JBM_employeemaster where emplogin=pm.ProjectCoordInd) as PM,ji.BM_DesiredPgCount as PgCount,ji.JBM_Intrnl,(ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)% 3 AS Col, (ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)/ 3 AS Row FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and   pm.current_Status is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                // string strQueryFinal = "Select ji.JBM_AutoID, jc.CustSN, ji.JBM_ID, ji.JBM_Intrnl,(ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)%3 AS Col, (ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)/3 AS Row from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";

                if (sCustSN != "AllCustomers")
                {
                    strQueryFinal += " and jc.CustSN in ('" + sCustSN.Trim() + "')";
                }

                if (sServices != "AllServices")
                {
                    strQueryFinal += " and ji.BM_FullService = '" + sServices.Trim() + "'";
                }

                if (sStatus != "AllStatus")
                {
                    if (sStatus != "DateRange")
                    {
                        if (sStatus != "Live Projects")
                        {
                            if (strrStatus != "" && strrStatus != null)
                            {
                                strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                            }
                            else
                                strQueryFinal += " and pm.Current_Status = '" + sStatus.Trim() + "'";
                        }
                        else
                        {
                            if (strrStatus != "" && strrStatus != null)
                            {
                                strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                            }
                            else
                                strQueryFinal += " and pm.Current_Status in('In Progress','On Hold','Yet to Start')";
                        }
                    }
                    else
                    {
                        if (strrStatus != "" && strrStatus != null)
                        {
                            strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                        }
                        strQueryFinal += " and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + Session["sStartDate"].ToString() + "', 101)  and CONVERT(DATETIME, '" + Session["sEndDate"].ToString() + "',101)";
                    }
                }
                else
                {
                    if (strrStatus != "" && strrStatus != null)
                    {
                        strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                    }
                }

                if (sHealth != "AllHealth")
                {
                    strQueryFinal += " and pm.Current_Health = '" + sHealth.Trim() + "'";
                }

                if (sCPM != "AllCPMs")
                {
                    strQueryFinal += " and pm.ProjectManagerUS = '" + sCPM.Trim() + "'";
                }

                if (sPMProject == "MyProjects" && Session["CustomerSN"].ToString() == "")   //LPM/PM
                {
                    strQueryFinal += " and pm.ProjectCoordInd = '" + Session["EmpLogin"].ToString().Trim() + "'";
                }
                else if (sPMProject == "MyTeam") //PM/LPM
                {
                    //strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "'  and  pm.ProjectCoordInd in (select emplogin from JBM_Employeemaster where usergroup=(select emplogin from JBM_Employeemaster where emplogin='" + Session["EmpLogin"].ToString().Trim() + "' and roleid='103') and roleid='104') ";
                    strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' ";
                }

                if (!Regex.IsMatch(sPMProject, "(MyProjects|MyTeam|AllProjects|AllPMs)", RegexOptions.IgnoreCase))
                // if (sPMProject != "MyProjects" && sPMProject != "MyTeam" && sPMProject != "AllProjects" && sPMProject != "AllPMs")
                {
                    //strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                    strQueryFinal += "  and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                }
                if (sDivision != "AllDivisions")
                {
                    if (sDivision == "N/A")
                        sDivision = "0";
                    else if (sDivision == "HSS")
                        sDivision = "1";
                    else if (sDivision == "STEM")
                        sDivision = "2";
                    strQueryFinal += " and ji.JBM_Location = '" + sDivision.Trim() + "'";
                }
                if (Session["sesSite"].ToString().Trim() != "")
                {
                    strQueryFinal += " and pm.Cenveo_Facility = '" + Session["sesSite"].ToString().Trim() + "'";
                }
                if (strHealth != "" && strHealth != null)
                {
                    strQueryFinal += " and pm.Current_Health = '" + strHealth.Trim() + "'";
                }


                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by JBM_ID asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLink(a[0].ToString(),"HrefLink", a[5].ToString(), a[13].ToString()),
                                     a[5].ToString(),
                                     a[1].ToString(),
                                     a[8].ToString(),
                                     a[10].ToString(),
                                     a[11].ToString().Trim()=="[Select]"?"":a[11].ToString().Trim(),
                                     a[12].ToString(),
                                     a[9].ToString().Trim()!=""? a[9].ToString().Trim()=="0"?"Full Service":"Comp":"",
                                     a[3].ToString(),
                                     HighlightHealth(a[2].ToString().Trim(), "Projects"),
                                     dt_DateFrmtSort(a[6].ToString()) + a[6].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult GetProjectListforReports(string sCustSN, string sCPM, string sPMProject, string sServices, string sStatus, string sHealth,string sDivision,string sSite,string strHealth, string strrStatus)
        {
            try
            {
                Session["sPMProject"] = sPMProject;
                ViewBag.strdashboard = "";
                ViewBag.ServiceID = "";
                ViewBag.uniqueName = "";
                ViewBag.PMID = "";
                ViewBag.CPMID = "";
                ViewBag.SiteID = "";
                ViewBag.DivID = "";
                ViewBag.sStatus = "";
                ViewBag.sCstSN = "";
                ViewBag.sHealth = "";
                ViewBag.strStatus = "";

                string strRoleID = Session["RoleID"].ToString();
                if (Session["CustomerSN"].ToString() != "")
                {
                    sCustSN = Session["CustomerSN"].ToString();
                }
                    
                //Session assign for to filter schedule and remarks tab
                Session["sesCustSN"] = sCustSN;
                Session["sesCPM"] = sCPM;
                Session["sesPMProject"] = sPMProject;
                Session["sesServices"] = sServices;
                
                if (strHealth!="" && strHealth!= null)
                    Session["sesHealth"] = strHealth;
                else
                    Session["sesHealth"] = sHealth;
                if (strrStatus != "" && strrStatus != null)
                    Session["sesStatus"] = strrStatus;
                else
                    Session["sesStatus"] = sStatus;

                string strQueryFinal = "SELECT ji.JBM_AutoID,CASE WHEN ji.JBM_Location='0' THEN 'N/A' WHEN ji.JBM_Location='1' THEN 'HSS' WHEN ji.JBM_Location='2' THEN 'STEM' ELSE ji.JBM_Location END  As Division, pm.Current_Health, pm.Current_Status, jc.CustName, jc.CustSN, FORMAT(CAST(ji.JBM_PrinterDate as date), 'dd-MMM-yy') as JBM_PrinterDate, ji.JBM_ID, ji.Title, ji.BM_FullService, pm.ProjectManagerUS as CPM, (select empname from JBM_employeemaster where emplogin=pm.ProjectCoordInd) as PM,(case when (select sum (case when [ActualPages] is null then 0 else [ActualPages] end) as [ActualPages] From  BK_ChapterInfo WHERE JBM_AutoID=ji.JBM_AutoID) = 0 then (select BM_DesiredPgCount from JBM_Info where JBM_AutoID=ji.JBM_AutoID) else (select sum(case when[ActualPages] is null then 0 else [ActualPages] end) as [ActualPages] From BK_ChapterInfo WHERE JBM_AutoID = ji.JBM_AutoID) end) as [PgCount],ji.JBM_Intrnl,(ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)% 3 AS Col, (ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)/ 3 AS Row FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and   pm.current_Status is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                //string strQueryFinal = "SELECT ji.JBM_AutoID,CASE WHEN ji.JBM_Location='0' THEN 'N/A' WHEN ji.JBM_Location='1' THEN 'HSS' WHEN ji.JBM_Location='2' THEN 'STEM' ELSE ji.JBM_Location END  As Division, pm.Current_Health, pm.Current_Status, jc.CustName, jc.CustSN, FORMAT(CAST(ji.JBM_PrinterDate as date), 'dd-MMM-yy') as JBM_PrinterDate, ji.JBM_ID, ji.Title, ji.BM_FullService, pm.ProjectManagerUS as CPM, (select empname from JBM_employeemaster where emplogin=pm.ProjectCoordInd) as PM,ji.BM_DesiredPgCount as PgCount,ji.JBM_Intrnl,(ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)% 3 AS Col, (ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)/ 3 AS Row FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and   pm.current_Status is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                // string strQueryFinal = "Select ji.JBM_AutoID, jc.CustSN, ji.JBM_ID, ji.JBM_Intrnl,(ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)%3 AS Col, (ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)/3 AS Row from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";

                if (sCustSN != "AllCustomers")
                {
                    strQueryFinal += " and jc.CustSN in ('" + sCustSN.Trim() + "')";
                }

                if (sServices != "AllServices")
                {
                    strQueryFinal += " and ji.BM_FullService = '" + sServices.Trim() + "'";
                }

                if (sStatus != "AllStatus")
                {
                    if (sStatus != "DateRange")
                    {
                        if (sStatus != "Live Projects")
                        {
                            if (strrStatus != "" && strrStatus != null)
                            {
                                strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                            }
                            else
                                strQueryFinal += " and pm.Current_Status = '" + sStatus.Trim() + "'";
                        }
                        else
                        {
                            if (strrStatus != "" && strrStatus != null)
                            {
                                strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                            }
                            else
                                strQueryFinal += " and pm.Current_Status in('In Progress','On Hold','Yet to Start')";
                        }
                    }
                    else
                    {
                        if (strrStatus != "" && strrStatus != null)
                        {
                            strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                        }
                        strQueryFinal += " and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + Session["sStartDate"].ToString() + "', 101)  and CONVERT(DATETIME, '" + Session["sEndDate"].ToString() + "',101)";
                    }
                }
                else
                {
                    if (strrStatus != "" && strrStatus != null)
                    {
                        strQueryFinal += " and pm.Current_Status = '" + strrStatus.Trim() + "'";
                    }
                }
                if (sHealth != "AllHealth")
                {
                    strQueryFinal += " and pm.Current_Health = '" + sHealth.Trim() + "'";
                }

                if (sCPM != "AllCPMs")
                {
                    strQueryFinal += " and pm.ProjectManagerUS = '" + sCPM.Trim() + "'";
                }

                if (sPMProject == "MyProjects" && Session["CustomerSN"].ToString() == "")   //LPM/PM
                {
                    strQueryFinal += " and pm.ProjectCoordInd = '" + Session["EmpLogin"].ToString().Trim() + "'";
                }
                else if (sPMProject == "MyTeam") //PM/LPM
                {
                    strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' ";
                    //strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "'  and  pm.ProjectCoordInd in (select emplogin from JBM_Employeemaster where usergroup=(select emplogin from JBM_Employeemaster where emplogin='" + Session["EmpLogin"].ToString().Trim() + "' and roleid='103') and roleid='104') ";
                }

                if (!Regex.IsMatch(sPMProject, "(MyProjects|MyTeam|AllProjects|AllPMs)", RegexOptions.IgnoreCase))
               // if (sPMProject != "MyProjects" && sPMProject != "MyTeam" && sPMProject != "AllProjects" && sPMProject != "AllPMs")
                {
                    strQueryFinal += "  and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                    //strQueryFinal += "and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                }
                if(sDivision != "AllDivisions")
                {
                    if (sDivision == "N/A")
                        sDivision = "0";
                    else if(sDivision == "HSS")
                        sDivision = "1";
                    else if (sDivision == "STEM")
                        sDivision = "2";
                    strQueryFinal += " and ji.JBM_Location = '" + sDivision.Trim() + "'";
                }
                if (Session["sesSite"].ToString().Trim() != "")
                {
                    strQueryFinal += " and pm.Cenveo_Facility = '" + Session["sesSite"].ToString().Trim() + "'";
                }
                if (strHealth != "" && strHealth!=null)
                {
                    strQueryFinal += " and pm.Current_Health = '" + strHealth.Trim() + "'";
                }
                

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by JBM_ID asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                  "<input type='checkbox' class='caseChk' id='" + a[0].ToString() + "' onclick=\"funcCheckItem('" + a[0].ToString() + "')\" class='form-control text-center' name='" + a[0].ToString() + "' value='KGLs'>",
                                     "<label class='control-label' for= '" + a[0].ToString() + "'>"+a[13].ToString()+"</ label > ",
                                     a[5].ToString(),
                                     a[1].ToString(),
                                     a[8].ToString(),
                                     a[10].ToString(),
                                     a[11].ToString().Trim()=="[Select]"?"":a[11].ToString().Trim(),
                                    
                                     a[9].ToString().Trim()!=""? a[9].ToString().Trim()=="0"?"Full Service":"Comp":"",
                                     a[3].ToString(),
                                     HighlightHealth(a[2].ToString().Trim(), "Projects")
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        [DeleteFileAttribute] //Action Filter, it will auto delete the file after download, I will explain it later
        public ActionResult Download(string file)
        {
            //get the temp folder and file path in server
            string fullPath = Path.Combine(Server.MapPath("~/UploadedFiles"), file);
            //return the file for download, this is an Excel so I set the file content type to "application/vnd.ms-excel"
            return File(fullPath, "application/vnd.ms-excel", file);
        }
        public class DeleteFileAttribute : ActionFilterAttribute
        {
            public override void OnResultExecuted(ResultExecutedContext filterContext)
            {
                filterContext.HttpContext.Response.Flush();
                //convert the current filter context to file and get the file path
                string filePath = (filterContext.Result as FilePathResult).FileName;

                //delete the file after download
                System.IO.File.Delete(filePath);
            }
        }
        [SessionExpire]
        public ActionResult ProjectExport(string SaveItemColloc) 
        {
            try
            {
                DataSet dsnew = new DataSet();
                string strQueryPrimarWithOther = "";
                DataTable dataTable = new DataTable();
                DataSet dsschedule = new DataSet();
                string JBM_AutoIDs = "";
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(SaveItemColloc);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string sAutoArtID = "";
                        string sCustomer = "";
                        Session["CustomerName"] = "";
                        if (saveIds[i].Contains("|"))
                        {
                            sAutoArtID = saveIds[i].Split('|')[0];
                            sCustomer = saveIds[i].Split('|')[1];

                            DataSet ds1 = new DataSet();
                            ds1 = DBProc.GetResultasDataSet(@"Select CustName from JBM_CustomerMaster WHERE CustSN='" + sCustomer + "'", Session["sConnSiteDB"].ToString());
                            Session["CustomerName"] = ds1.Tables[0].Rows[0][0].ToString();

                        }
                        else { sAutoArtID = saveIds[i]; }
                        
                        JBM_AutoIDs += "'"+sAutoArtID+"',";
                        DataSet dsstage = new DataSet();
                        dsstage = DBProc.GetResultasDataSet(@"SELECT StageID FROM   BK_ProjectManagement where JBM_AutoID ='" + sAutoArtID + "'", Session["sConnSiteDB"].ToString());
                        string stageID = dsstage.Tables[0].Rows[0][0].ToString();
                        if (stageID != "")
                            stageID = "%" + stageID + "%";
                        else
                            stageID = "";
                        strQueryPrimarWithOther = @"Select SD.StageName as [PrimaryStages], CASE WHEN SI.Short_Stage = 'CE' THEN 
(select  max(CEDueDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " +
"" + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + sAutoArtID + "' and " +
"(Active is null or Active=1)) and RevFinStage='FP') ELSE (select  max(DueDate) from " + Session["sCustAcc"].ToString() +
"_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='"
+sAutoArtID + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [Due], " +
" CASE WHEN SI.Short_Stage = 'CE' THEN (select  max(CERevisedDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where " +
"AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" +
sAutoArtID + "' and (Active is null or Active=1)) and RevFinStage='FP') ELSE (select  max(RevisedDate) from "
+ Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() +
"_ChapterInfo where JBM_AutoID='" + sAutoArtID + "' and (Active is null or Active=1)) and " +
"RevFinStage=SI.Short_Stage) END AS [Revised], CASE WHEN SI.Short_Stage = 'CE' THEN (select  max(CeDispDate) from "
+ Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() +
"_ChapterInfo where JBM_AutoID='" + sAutoArtID + "' and (Active is null or Active=1)) and RevFinStage='FP') " +
"WHEN SI.Short_Stage = '1pco' THEN ((SELECT MAX(CorrMaxDate) AS LastUpdateDate FROM (select PrintFinalDue  AS CorrMaxDate from " +
Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() +
"_ChapterInfo where JBM_AutoID='" + sAutoArtID + "' and (Active is null or Active=1)) and PrintFinalDue" +
" is not null union all select  PR_corr_Appr  AS CorrMaxDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID " +
"in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + sAutoArtID
+ "' and (Active is null or Active=1)) and PR_corr_Appr is not null union all select  Aut_Corr_Appr  AS CorrMaxDate from " +
Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() +
"_ChapterInfo where JBM_AutoID='" + sAutoArtID + "' and (Active is null or Active=1)) and Aut_Corr_Appr " +
"is not null) CorrectionDate)) ELSE (select  max(DispatchDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID" +
" in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + sAutoArtID
+ "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [Actual], SI.SeqID from " +
Session["sCustAcc"].ToString() + "_ScheduleInfo SI JOIN JBM_StageDescription SD ON SI.Short_Stage = SD.StageShortName  " +
"WHERE SI.JBM_AutoID='" + sAutoArtID + "' and SD.Is_CustStage = 'Y' and SD.CustType='" + sCustomer + "' and" +
" CustStageID like '" + Session["StageID"] + "' and SD.StageSeqID in (1,2,3,4,5,6,7,8,9,10,11,12,13)  and SI.DeleteYN ='N' or " +
"SI.DeleteYN is null and JBM_AutoID='" + sAutoArtID + "' and SD.CustType='" + sCustomer + "' and CustStageID like '"
+ stageID + "' and SD.StageSeqID < 14 ";

                        // To add other stages information from ScheduleInfo tabel
                        strQueryPrimarWithOther += " UNION Select SD.StageName as [PrimaryStages],PlannedEndDate as Due,RevisedPlanEndDate as Revised,ActualEndDate as Actual,SeqID from " + Session["sCustAcc"].ToString() + "_ScheduleInfo SI JOIN JBM_StageDescription SD ON SI.Short_Stage = SD.StageShortName  where  SI.SeqID > 13 and JBM_AutoID='" + sAutoArtID + "' and SD.CustType='" + sCustomer + "' and CustStageID like '" + stageID + "' order by SeqID"; //SI.DeleteYN ='N' or SI.DeleteYN is null and

                       DataSet dsSchedulenew = DBProc.GetResultasDataSet(strQueryPrimarWithOther, Session["sConnSiteDB"].ToString());
                        if (dsSchedulenew.Tables.Count > 0)
                        {
                            DataTable dtCopy = dsSchedulenew.Tables[0].Copy();
                            dtCopy.TableName = sAutoArtID;
                            dsschedule.Tables.Add(dtCopy);
                        }
                    }
                    JBM_AutoIDs=JBM_AutoIDs.Remove(JBM_AutoIDs.Length - 1, 1);

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(@"SELECT JI.JBM_AutoID,Title,BM_Author as Author,BM_ISBN10number as ISBN ,
  CASE
                        WHEN PM.ProjectCoordInd = '[Select]' THEN ''
                        ELSE PM.ProjectCoordInd END as PM,
PM.Copyeditor AS Editor, PM.TeamLead, 
                        CASE
                           WHEN JI.JBM_Location = 1 THEN 'HSS'
                           WHEN JI.JBM_Location = 2 THEN 'STEM'
                        ELSE '' END as Section,CASE
                           WHEN JI.JBM_Trimsize = '[Select]' THEN ''
                        ELSE JI.JBM_Trimsize END as TrimSize,'' as EstimatedExtent, OverallCost, AdditionalCost
                        FROM   BK_ProjectManagement PM INNER JOIN JBM_Info JI ON PM.JBM_AutoID = JI.JBM_AutoID where JI.JBM_AutoID in(" + JBM_AutoIDs + ") ", Session["sConnSiteDB"].ToString());
                    dataTable = ds.Tables[0];                                        

                  
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {                        
                        DataTable dataTablenew = new DataTable();
                        dataTablenew.Columns.Add("JBM_AutoID", typeof(String));
                        dataTablenew.Columns.Add("Title", typeof(String));
                        dataTablenew.Columns.Add("Author", typeof(String));
                        dataTablenew.Columns.Add("ISBN", typeof(String));
                        dataTablenew.Columns.Add("PM", typeof(String));
                        dataTablenew.Columns.Add("Editor", typeof(String));
                        dataTablenew.Columns.Add("TeamLead", typeof(String));
                        dataTablenew.Columns.Add("Section", typeof(String));
                        dataTablenew.Columns.Add("TrimSize", typeof(String));
                        dataTablenew.Columns.Add("EstimatedExtent", typeof(String));
                        dataTablenew.Columns.Add("OverallCost", typeof(String));
                        dataTablenew.Columns.Add("AdditionalCost", typeof(String));
                       
                        DataRow dr = dataTable.Rows[i];
                        dataTablenew.ImportRow(dr);
                        DataRow workRow = dataTablenew.NewRow();
                        dataTablenew.Rows.InsertAt(workRow, dataTablenew.Rows.Count);
                        DataTable dtnew = new DataTable();
                        if (dsschedule.Tables.Count > 0)
                        {
                            dtnew = dsschedule.Tables[dataTable.Rows[i]["JBM_AutoID"].ToString()].Copy();
                        }
                        workRow = dtnew.NewRow();
                        dtnew.Rows.InsertAt(workRow, dtnew.Rows.Count);
                        DataTable dtnotes = new DataTable();
                        dtnotes = DBProc.GetResultasDataTbl(@"Select Instruction,CAST(InstDate as VARCHAR(50)) as Date from BK_SplInstructions 
                    where AutoArtID='" + dataTable.Rows[i]["JBM_AutoID"].ToString() + "' ORDER BY CONVERT(DateTime, InstDate,101)  DESC", Session["sConnSiteDB"].ToString());
                        workRow = dtnotes.NewRow();
                        dtnotes.Rows.InsertAt(workRow, dtnotes.Rows.Count);
                        dataTablenew.TableName = "Project"+i;
                        dtnotes.TableName = "Note"+i;

                        dsnew.Tables.Add(dataTablenew);
                        dsnew.Tables.Add(dtnew);
                        dsnew.Tables.Add(dtnotes);
                        
                        //ToCSV(dataTablenew, "D:\\Report_" + JBM_AutoIDs + ".csv");
                        //ToCSV(dtnew,"D:\\Report_" + JBM_AutoIDs + ".csv");
                        //ToCSV(dtnotes, "D:\\Report_" + JBM_AutoIDs + ".csv");
                    }
                    var datetime = DateTime.Now.ToString("MMddyyyyHHmmss");
                                       
                    string fileName = "Report_" + datetime + ".xlsx";
                    CreateExcelFile(dsnew, Server.MapPath("~/UploadedFiles/Report_" + datetime + ".xlsx"),fileName);
                       //"D:\\Report_" + JBM_AutoIDs + ".xlsx"

                    //old excel code

                    //                   Microsoft.Office.Interop.Excel.Application excel;
                    //                   Microsoft.Office.Interop.Excel.Workbook excelworkBook;
                    //                   Microsoft.Office.Interop.Excel.Worksheet excelSheet;
                    //                   // Start Excel and get Application object.  
                    //                   excel = new Microsoft.Office.Interop.Excel.Application();
                    //                   // for making Excel visible  
                    //                   excel.Visible = false;
                    //                   excel.DisplayAlerts = false;
                    //                   // Creation a new Workbook  
                    //                   excelworkBook = excel.Workbooks.Add(Type.Missing);
                    //                   // Workk sheet  
                    //                   excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    //                   excelSheet.Name = "Sheet";

                    //                   for (int i = 2; i <= 5; i++) // this will apply it from col 1 to 10
                    //                   {
                    //                       excelSheet.Columns[i].ColumnWidth = 30;
                    //                   }
                    //                   excelSheet.Columns[7].ColumnWidth = 20;
                    //                   excelSheet.Columns[8].ColumnWidth = 50;
                    //                   Excel.Range chartRange;    
                    //                   excelSheet.get_Range("a1", "l1").Merge(false);
                    //                   chartRange = excelSheet.get_Range("a1", "l1");
                    //                   chartRange.EntireRow.Font.Bold = true;
                    //                   chartRange.FormulaR1C1 = "Project Management Report for Taylor & Francis!";
                    //                   chartRange.HorizontalAlignment = 3;
                    //                   chartRange.VerticalAlignment = 3;
                    //                   chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                    //                   chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                    //                   int rowstart = 2;
                    //                   for (int i=0;i<dataTable.Rows.Count;i++)
                    //                   {
                    //                       DataTable dtnew = new DataTable();
                    //                       if (dsschedule.Tables.Count>0)
                    //                       {                         
                    //                           dtnew = dsschedule.Tables[dataTable.Rows[i]["JBM_AutoID"].ToString()].Copy();
                    //                       }
                    //                       int rowsize = 10;
                    //                       if(dtnew.Rows.Count>0)
                    //                       {
                    //                           rowsize = rowsize+ dtnew.Rows.Count+2;
                    //                       }
                    //                       rowstart = rowstart + 1;
                    //                       int newrowsize = rowstart + rowsize;
                    //                       Excel.Range formatRange;
                    //                       formatRange = excelSheet.get_Range("b"+ rowstart, "e"+ newrowsize);
                    //                       formatRange.BorderAround(Excel.XlLineStyle.xlContinuous,
                    //                       Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic,
                    //                       Excel.XlColorIndex.xlColorIndexAutomatic);

                    //                       formatRange = excelSheet.get_Range("g" + rowstart, "h" + newrowsize);
                    //                       formatRange.BorderAround(Excel.XlLineStyle.xlContinuous,
                    //                       Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic,
                    //                       Excel.XlColorIndex.xlColorIndexAutomatic);

                    //                       DataTable dtnotes = new DataTable();
                    //                       dtnotes = DBProc.GetResultasDataTbl(@"Select Instruction,CAST(InstDate as VARCHAR(50)) as DateToString from BK_SplInstructions 
                    //where AutoArtID='"+ dataTable.Rows[i]["JBM_AutoID"].ToString() + "' ORDER BY CONVERT(DateTime, InstDate,101)  DESC", Session["sConnSiteDB"].ToString());

                    //                       int y = 0;
                    //                       int startnotes = 0;
                    //                       int start = 0;
                    //                       int x = 0;
                    //                       for(int j= rowstart;j<= newrowsize;j++)
                    //                       {
                    //                           if (start < 3)
                    //                           {
                    //                               string columnName = dataTable.Columns[start].ColumnName;
                    //                               formatRange = excelSheet.get_Range("b" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               excelSheet.get_Range("c" + j, "e" + j).Merge(false);
                    //                               formatRange = excelSheet.get_Range("c" + j, "e" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               start++;
                    //                           }
                    //                           else if (start == 3 || start == 6)
                    //                               start++;
                    //                           else if (start == 4)
                    //                           {
                    //                               string columnName = dataTable.Columns[start - 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("b" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start].ColumnName;
                    //                               formatRange = excelSheet.get_Range("c" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("d" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 2].ColumnName;
                    //                               formatRange = excelSheet.get_Range("e" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               start++;
                    //                           }
                    //                           else if (start == 5)
                    //                           {
                    //                               string columnName = dataTable.Columns[start - 2].ColumnName;
                    //                               formatRange = excelSheet.get_Range("b" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start - 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("c" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start].ColumnName;
                    //                               formatRange = excelSheet.get_Range("d" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("e" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               start++;
                    //                           }
                    //                           else if (start == 7)
                    //                           {
                    //                               string columnName = dataTable.Columns[start].ColumnName;
                    //                               formatRange = excelSheet.get_Range("b" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("c" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 2].ColumnName;
                    //                               formatRange = excelSheet.get_Range("d" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 3].ColumnName;
                    //                               formatRange = excelSheet.get_Range("e" + j);
                    //                               formatRange.FormulaR1C1 = columnName;
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               start++;
                    //                           }
                    //                           else if (start == 8)
                    //                           {
                    //                               string columnName = dataTable.Columns[start - 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("b" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start].ColumnName;
                    //                               formatRange = excelSheet.get_Range("c" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 1].ColumnName;
                    //                               formatRange = excelSheet.get_Range("d" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               columnName = dataTable.Columns[start + 2].ColumnName;
                    //                               formatRange = excelSheet.get_Range("e" + j);
                    //                               formatRange.FormulaR1C1 = dataTable.Rows[i][columnName].ToString();
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               start++;
                    //                           }
                    //                           else if (start == 9)
                    //                               start++;
                    //                           else if (start == 10)
                    //                           {
                    //                               excelSheet.get_Range("b" + j, "e" + j).Merge(false);
                    //                               formatRange = excelSheet.get_Range("b" + j, "e" + j);
                    //                               formatRange.FormulaR1C1 = "SCHEDULE";
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               start++;
                    //                           }
                    //                           else if (start == 11)
                    //                           {
                    //                               if (dsschedule.Tables.Count > 0)
                    //                               {
                    //                                   dtnew = dsschedule.Tables[dataTable.Rows[i]["JBM_AutoID"].ToString()].Copy();
                    //                                   if (dtnew.Rows.Count > 0)
                    //                                   {
                    //                                       formatRange = excelSheet.get_Range("b" + j);
                    //                                       formatRange.FormulaR1C1 = "Action";
                    //                                       formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                       formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                       formatRange = excelSheet.get_Range("c" + j);
                    //                                       formatRange.FormulaR1C1 = "Due";
                    //                                       formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                       formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                       formatRange = excelSheet.get_Range("d" + j);
                    //                                       formatRange.FormulaR1C1 = "Revised";
                    //                                       formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                       formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                       formatRange = excelSheet.get_Range("e" + j);
                    //                                       formatRange.FormulaR1C1 = "Actual";
                    //                                       formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                       formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                                       start++;
                    //                                   }
                    //                               }

                    //                           }
                    //                           else if (start > 11)
                    //                           {
                    //                               if (dsschedule.Tables.Count > 0)
                    //                               {
                    //                                   dtnew = dsschedule.Tables[dataTable.Rows[i]["JBM_AutoID"].ToString()].Copy();
                    //                                   if (dtnew.Rows.Count > 0)
                    //                                   {
                    //                                       if (dtnew.Rows.Count > x)
                    //                                       {
                    //                                           string columnName = dtnew.Columns[0].ColumnName;
                    //                                           formatRange = excelSheet.get_Range("b" + j);
                    //                                           formatRange.FormulaR1C1 = dtnew.Rows[x][columnName].ToString();
                    //                                           formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                           formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                           columnName = dtnew.Columns[1].ColumnName;
                    //                                           formatRange = excelSheet.get_Range("c" + j);
                    //                                           formatRange.FormulaR1C1 = dtnew.Rows[x][columnName].ToString();
                    //                                           formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                           formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                           columnName = dtnew.Columns[2].ColumnName;
                    //                                           formatRange = excelSheet.get_Range("d" + j);
                    //                                           formatRange.FormulaR1C1 = dtnew.Rows[x][columnName].ToString();
                    //                                           formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                           formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                           columnName = dtnew.Columns[3].ColumnName;
                    //                                           formatRange = excelSheet.get_Range("e" + j);
                    //                                           formatRange.FormulaR1C1 = dtnew.Rows[x][columnName].ToString();
                    //                                           formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                                           formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                                           x++;
                    //                                           start++;
                    //                                       }
                    //                                   }
                    //                               }
                    //                           }

                    //                           if (startnotes == 0)
                    //                           {
                    //                               excelSheet.get_Range("g" + j, "h" + j).Merge(false);
                    //                               formatRange = excelSheet.get_Range("g" + j, "h" + j);
                    //                               formatRange.FormulaR1C1 = "COMMENTS DETAILS";
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               startnotes++;
                    //                           }
                    //                           else if (startnotes == 1)
                    //                           {
                    //                               formatRange = excelSheet.get_Range("g" + j);
                    //                               formatRange.FormulaR1C1 = "Date";
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                               formatRange = excelSheet.get_Range("h" + j);
                    //                               formatRange.FormulaR1C1 = "Comments";
                    //                               formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    //                               formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //                               startnotes++;
                    //                           }
                    //                           else if (startnotes > 1)
                    //                           {
                    //                               if (dtnotes.Rows.Count > y)
                    //                               {
                    //                                   string columnName = dtnotes.Columns[0].ColumnName;
                    //                                   formatRange = excelSheet.get_Range("h" + j);
                    //                                   formatRange.FormulaR1C1 = dtnotes.Rows[y][columnName].ToString();
                    //                                   formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    //                                   formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                   columnName = dtnotes.Columns[1].ColumnName;
                    //                                   formatRange = excelSheet.get_Range("g" + j);
                    //                                   formatRange.FormulaR1C1 = dtnotes.Rows[y][columnName].ToString();
                    //                                   formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    //                                   formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    //                                   startnotes++;
                    //                                   y++;
                    //                               }
                    //                           }
                    //                       }


                    //                       rowstart = newrowsize + 1;
                    //                   }

                    //                   //now save the workbook and exit Excel
                    //                   excelworkBook.SaveAs("D:\\Report_"+ JBM_AutoIDs + ".xlsx"); ;
                    //                   excelworkBook.Close();
                    //                   excel.Quit();

                    return Json(new { dataComp =fileName }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        public void CreateExcelFile(DataSet data, string OutPutFileDirectory, string fileName)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(OutPutFileDirectory, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }           
        }

        private void CreatePartsForExcel(SpreadsheetDocument document, DataSet data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);            

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");            
            GenerateWorksheetPartContent(worksheetPart1, partSheetData, data);           
            
        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1, DataSet data)
        {
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet.Append(sheetDimension1);
            worksheet.Append(sheetViews1);
            worksheet.Append(sheetFormatProperties1);
            worksheet.Append(sheetData1);
            worksheet.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet;
            int count = 1;
            MergeTwoCells(worksheet, "A" + count, "K" + count);
            count++;
            for (int i = 0; i < data.Tables.Count; i++)
            {
                MergeTwoCells(worksheet, "A" + count, "K" + count);
                count = count + data.Tables[i].Rows.Count + 2;
            }
            SetColumnWidth(worksheet, 1U, 60);
        }
        static void SetColumnWidth(Worksheet worksheet, uint Index, DoubleValue dwidth)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns cs = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (cs != null)
            {
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Column> ic = cs.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(r => r.Min == Index).Where(r => r.Max == Index);
                if (ic.Count() > 0)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = ic.First();
                    c.Width = dwidth;
                }
                else
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                    cs.Append(c);
                }
            }
            else
            {
                cs = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                cs.Append(c);
                worksheet.InsertAfter(cs, worksheet.GetFirstChild<SheetFormatProperties>());
            }
        }
        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)3U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            // FillId = 2,RED
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "93ebf9" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
            
        }
        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        private SheetData GenerateSheetdataForDetails(DataSet data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateMainHeaderRowForExcel("Main",""));
            for (int i = 0; i < data.Tables.Count; i++)
            {
                string tablename = data.Tables[i].TableName;
                sheetData1.Append(CreateMainHeaderRowForExcel(tablename, data.Tables[i].Rows[0][0].ToString()));
                sheetData1.Append(CreateHeaderRowForExcel(tablename));
                foreach (DataRow taktTimemodel in data.Tables[i].Rows)
                {
                    Row partsRows = GenerateRowForChildPartDetail(taktTimemodel, tablename);
                    sheetData1.Append(partsRows);
                }
            }

            return sheetData1;
        }
        private Row CreateHeaderRowForExcel(string tablename)
        {
            Row workRow = new Row();
            if(tablename.Contains("Project"))
            { tablename = "Project"; }
            else if (tablename.Contains("Note"))
            { tablename = "Note"; }
            if (tablename == "Project")
            {
                workRow.Append(CreateCell("Title"));
                workRow.Append(CreateCell("Author"));
                workRow.Append(CreateCell("ISBN"));
                workRow.Append(CreateCell("PM"));
                workRow.Append(CreateCell("Editor"));
                workRow.Append(CreateCell("Team Lead"));
                workRow.Append(CreateCell("Section"));
                workRow.Append(CreateCell("Trim Size"));
                workRow.Append(CreateCell("Estimated Extent"));
                workRow.Append(CreateCell("Overall Cost"));
                workRow.Append(CreateCell("Additional Cost"));
            }
            else if (tablename == "Note")
            {
                workRow.Append(CreateCell("Instruction"));
                workRow.Append(CreateCell("Date"));
            }
            else
            {
                workRow.Append(CreateCell("Action"));
                workRow.Append(CreateCell("Due"));
                workRow.Append(CreateCell("Revised"));
                workRow.Append(CreateCell("Actual"));
            }
            return workRow;
        }
        private Row CreateMainHeaderRowForExcel(string tablename,string ProjectID)
        {
            Row workRow = new Row();
            if (tablename.Contains("Project"))
            { tablename = "Project"; }
            else if (tablename.Contains("Note"))
            { tablename = "Note"; }
            if (tablename == "Project")
            {
                workRow.Append(CreateCell("PROJECT DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Note")
            {
                workRow.Append(CreateCell("COMMENTS DETAILS ", 1U, workRow.RowIndex));
            }
            else if (tablename == "Main")
            {
                workRow.Append(CreateCell("PROJECT MANAGEMENT REPORT FOR " + Session["CustomerName"].ToString(), 1U, workRow.RowIndex));
            }
            else
            {
                workRow.Append(CreateCell("SCHEDULE DETAILS", 1U, workRow.RowIndex));
            }
            return workRow;
        }
        private Row GenerateRowForChildPartDetail(DataRow testmodel,string tablename)
        {
            Row tRow = new Row();
            if (tablename.Contains("Project"))
            { tablename = "Project"; }
            else if (tablename.Contains("Note"))
            { tablename = "Note"; }
            if (tablename == "Project")
            {                
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[5].ToString()));
                tRow.Append(CreateCell(testmodel[6].ToString()));
                tRow.Append(CreateCell(testmodel[7].ToString()));
                tRow.Append(CreateCell(testmodel[8].ToString()));
                tRow.Append(CreateCell(testmodel[9].ToString()));
                tRow.Append(CreateCell(testmodel[10].ToString()));
                tRow.Append(CreateCell(testmodel[11].ToString()));
            }
            else if (tablename == "Note")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
            }
            else
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
            }
                return tRow;
        }
        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 2U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private Cell CreateCell(string text, uint styleIndex, UInt32Value RowIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
        private static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name)
        {

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.

            string s1 = cell1Name + ":" + cell2Name;
            MergeCell mergeCell = new MergeCell() { Reference = s1 };
            mergeCells.Append(mergeCell);

            worksheet.Save();

        }
        public void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, true);
            //headers
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
        //public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        //{
        //    range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
        //    range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
        //    if (IsFontbool == true)
        //    {
        //        range.Font.Bold = IsFontbool;
        //    }
        //}
        public static string dt_DateFrmtSort(string dtDateInput)
        {

            string strTemp = null;
            if (dtDateInput != "")
            {
                DateTime dateTime10 = Convert.ToDateTime(dtDateInput);
                strTemp = "<span style='display:none;'>" + dateTime10.ToString("yyyyMMdd").ToString() + "</span>";
            }
            return strTemp;
        }
        public string HighlightHealth(string strHealth, string  strTabType)
        {
            string formTag = string.Empty;
            try
            {
                strHealth = strHealth.Replace(" ", " ");

                if (strHealth == "On Track")
                {
                    if (strTabType == "Projects")
                    {
                        formTag = "<div style='background:#8edb18'><span style='display:none;'>4</span>" + strHealth + "</div>";
                    }
                    else {
                        formTag = "<div style='background:#8edb18;color:#8edb18;font-size:0px; width: 30px;height: 20px;'>4</div>";
                    }
                    
                }
                else if (strHealth == "In Escalation")
                {
                    if (strTabType == "Projects")
                    {
                        formTag = "<div style='background:#ff0909'><span style='display:none;'>2</span>" + strHealth + "</div>";
                    }
                    else {
                        formTag = "<div style='background:#ff0909;color:#ff0909;font-size:0px; width: 30px;height: 20px;'>2</div>";
                    }
                    
                }
                else if (strHealth == "N/A")
                {
                    if (strTabType == "Projects")
                    {
                        formTag = "<div style='background:#eeeeee'><span style='display:none;'>1</span>" + strHealth + "</div>";
                    }
                    else {
                        formTag = "<div style='background:#eeeeee;color:#eeeeee;font-size:0px; width: 30px;height: 20px;'>1</div>";
                    }
                    
                }
                else if (strHealth == "Watch Listed")
                {
                    if (strTabType == "Projects")
                    {
                        formTag = "<div style='background:#faff31'><span style='display:none;'>3</span>" + strHealth + "</div>";
                    }
                    else {

                        formTag = "<div style='background:#faff31;color:#faff31;font-size:0px; width: 30px;height: 20px;'>3</div>";
                    }
                    
                }
                else {
                    formTag = strHealth;
                }

                return formTag;
            }
            catch (Exception)
            {

              return "";
            }
        }
        public string CreateHREFLinkforDashboard(string sCustSN,string uniqueID, string strCustID,string strdisplay,string uniqueName,string strdashboard,string sStatus,string sHealth, string strStatus)
        {
            string formControl = string.Empty;
            try
            {
                if (strCustID != "")
                {
                    string strUrl = "";
                    strUrl = Request.Url.AbsoluteUri.ToString().Replace("GetCPMListbyHealth", "ProjectTracking" ).Replace("GetPMListbyHealth", "ProjectTracking").Replace("GetServiceListbyStatus", "ProjectTracking").Replace("GetServiceListbyHealth", "ProjectTracking").Replace("GetDivisionListbyHealth", "ProjectTracking").Replace("GetSiteListbyHealth", "ProjectTracking");
                  
                    string strNavigateURL = strUrl + "?sCustSN="+ sCustSN + "&CustId=" + strCustID + "&uniqueID=" + uniqueID + "&uniqueName=" + uniqueName+ "&strdashboard="+ strdashboard+ "&sStatus="+ sStatus+ "&sHealth="+ sHealth + "&strStatus=" + strStatus;
                    formControl = "<a href='" + strNavigateURL + "'>" + strdisplay + "</a>";
                }

                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        public string CreateHREFLink(string uniqueID, string strType, string strCustSN, string strJBMIntrnl)
        {
            string formControl = string.Empty;
            try
            {
                if (strCustSN != "")
                {
                    string strUrl = ""; 
                     strUrl = Request.Url.AbsoluteUri.ToString().Replace("ProjectTrack", "ProjectMgnt").Replace("GetProjectList", "Project").Replace("GetScheduleList", "Project").Replace("GetRemarksList", "Project");
                    GlobalVariables.strreturnURL = objDS.Encrypt(Session["returnURL"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strJBMAutoId = objDS.Encrypt(uniqueID, "*!%$@~&#?,:");
                    GlobalVariables.strCustSN = objDS.Encrypt(strCustSN, "*!%$@~&#?,:");
                    GlobalVariables.strEmpID = objDS.Encrypt(Session["EmpIdLogin"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strSiteID = objDS.Encrypt(Session["sSiteID"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strCustAcc = Session["sCustAcc"].ToString();

                    string strNavigateURL = strUrl + "?returnURL=" + GlobalVariables.strreturnURL + "&JBMAutoId=" + GlobalVariables.strJBMAutoId + "&CustSN=" + GlobalVariables.strCustSN + "&EmpID=" + GlobalVariables.strEmpID + "&SiteID=" + GlobalVariables.strSiteID + "&CustAcc=" + GlobalVariables.strCustAcc;
                    //sample : http://localhost:44358/ProjectMgnt/Project?returnURL=BZ1AAr3qrknHGTo5xB1Ue+2kZfivafi9yUc0/rq3uYJtA1/8t+0BfFy5cKCxz9tzcJwU+NjeuUAGdCC9rRvbuTpQvw4u9J0iKv9RWoye7oWPIh4c/aqclY10cVkPrpzulTsLpQqJHC4aMZqDLSrhfQ==&JBMAutoId=i4omwZDFxHE+Z245BgWtmA==&CustSN=5M5zCkmHL9kfBOZhnRXQAQ==&EmpID=1DyHJApHKbnzBP3CVO3/rg==&SiteID=focrQ+OtCm9E7Fy5buTZIg==&CustAcc=BK
                    formControl = "<a href='" + strNavigateURL + "'>" + strJBMIntrnl + "</a>";
                }



                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult Remarks()
        {

            try
            {
                if (Session["CustomerSN"].ToString() != "")
                {
                    string sCustSN = Session["CustomerSN"].ToString();
                    DataSet ds1 = new DataSet();
                    string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and jc.CustSN='" + sCustSN + "' order by CustSN asc ";
                    ds1 = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        ViewBag.vCustID = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        ViewBag.vCustSN = ds1.Tables[0].Rows[0]["CustSN"].ToString();
                        ViewBag.vCustName = ds1.Tables[0].Rows[0]["CustName"].ToString();
                        Session["CustomerName"] = ds1.Tables[0].Rows[0]["CustName"].ToString();
                    }
                }
                if (Session["RemarksPageLength"] == null)
                {
                    Session["RemarksPageLength"] = 10;
                }
                //Load Production Lead list Items
                List<SelectListItem> lstProductionLead = new List<SelectListItem>();
                DataSet ds = new DataSet();
                //ds = DBProc.GetResultasDataSet("Select DISTINCT ProductionLead from BK_ProjectManagement where ProductionLead is not null order by ProductionLead asc", Session["sConnSiteDB"].ToString() );
                ds = DBProc.GetResultasDataSet("Select EmpLogin,EmpName from jbm_employeemaster where roleid in ('103','104', '105') and emplogin not in ('20037','46109','kolstadb') order by EmpName asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        lstProductionLead.Add(new SelectListItem
                        {
                            Text = ds.Tables[0].Rows[intCount]["EmpName"].ToString(),
                            Value = ds.Tables[0].Rows[intCount]["EmpLogin"].ToString()
                        });
                    }
                }

                ViewBag.Productionlist = lstProductionLead;
                //Load Project Manager 
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                dsPM = DBProc.GetResultasDataSet("select EmpName,emplogin from JBM_Employeemaster where usergroup = (select emplogin from JBM_Employeemaster where emplogin = '" + Session["EmpLogin"].ToString().Trim() + "' and roleid = '103') and roleid = '104'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                        string stremplogin = dsPM.Tables[0].Rows[intCount]["emplogin"].ToString();

                        lstPM.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = stremplogin.ToString()
                        });
                    }

                }

                ViewBag.PMlist = lstPM;
                //Load Project Manager list Items
                List<SelectListItem> itemsProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpAutoID,EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='104' or roleid='103' order by empname asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.PrjMgrList = itemsProjectManager;
                ViewBag.PrjTabHead = "Projects";
                ViewBag.PageHead = "Project Tracking - Remarks";
            }
            catch (Exception)
            {
                return View();
            }
            return View();
        }
        [SessionExpire]
        public ActionResult GetServiceListbyStatus(string sStatus,string sStartDate,string sEndDate)
        {
            try//Project Status by Service
            {
                Session["sStartDate"] = sStartDate;
                Session["sEndDate"] = sEndDate;
                string strQueryFinal = "select BM_FullService as Service,projects,sum(pages) as pages,sum(ISNULL([In Progress], 0)) as [In progress],sum(ISNULL([On Hold], 0)) as [On Hold],sum(ISNULL([Yet to Start], 0)) as [Yet to Start], sum(ISNULL([Complete], 0)) as Complete,BM_FullServiceID from (SELECT pm.current_Status, sum(CAST(ji.BM_DesiredPgCount as int)) as pages, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.BM_FullService = ji.BM_FullService and  p.current_Status is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition'   WHEN ji.BM_FullService = 2 THEN 'End to End'END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Status is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Status,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Status IN([In Progress],[On Hold],[Yet to Start],[Complete])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                if (sStatus == "In Progress")
                {
                    strQueryFinal = "select BM_FullService as Service,projects,sum(pages) as pages,sum(ISNULL([In Progress], 0)) as [In progress],sum(ISNULL([On Hold], 0)) as [On Hold],sum(ISNULL([Yet to Start], 0)) as [Yet to Start], sum(ISNULL([Complete], 0)) as Complete,BM_FullServiceID from (SELECT pm.current_Status, sum(CAST(ji.BM_DesiredPgCount as int)) as pages, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold','Yet To Start') and j.BM_FullService = ji.BM_FullService and  p.current_Status is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition'   WHEN ji.BM_FullService = 2 THEN 'End to End'END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Status is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold','Yet To Start') group by pm.current_Status,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Status IN([In Progress],[On Hold],[Yet to Start],[Complete])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                }
                if (sStatus == "DateRange")
                {
                    strQueryFinal = "select BM_FullService as Service,projects,sum(pages) as pages,sum(ISNULL([In Progress], 0)) as [In progress],sum(ISNULL([On Hold], 0)) as [On Hold],sum(ISNULL([Yet to Start], 0)) as [Yet to Start], sum(ISNULL([Complete], 0)) as Complete,BM_FullServiceID from (SELECT pm.current_Status, sum(CAST(ji.BM_DesiredPgCount as int)) as pages, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '"+sStartDate+"', 101)  and CONVERT(DATETIME, '"+sEndDate+"',101) and j.BM_FullService = ji.BM_FullService and  p.current_Status is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition'   WHEN ji.BM_FullService = 2 THEN 'End to End'END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Status is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Status,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Status IN([In Progress],[On Hold],[Yet to Start],[Complete])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                }
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[0].ToString(), "Service","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[1].ToString(), "Service","DashBoard",sStatus,"",""):a[1].ToString(),
                                     a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[3].ToString(), "Service","DashBoard",sStatus,"","In Progress"):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[4].ToString(), "Service","DashBoard",sStatus,"","On Hold"):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[5].ToString(), "Service","DashBoard",sStatus,"","Yet to Start"):a[5].ToString(),
                                     (a[6].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[7].ToString(),Session["CustID"].ToString() , a[6].ToString(), "Service","DashBoard",sStatus,"","Complete"):a[6].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetServiceListbyHealth(string sStatus, string sStartDate, string sEndDate)
        {//Project Health by Service
            try
            {
                string strQueryFinal = "select BM_FullService as Service,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] ,BM_FullServiceID from   (SELECT pm.current_Health,  (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN   JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "'  and j.BM_FullService = ji.BM_FullService and  p.current_Health is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition' WHEN ji.BM_FullService = 2 THEN 'End to End' END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Health,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                if (sStatus == "In Progress")
                {
                    strQueryFinal = "select BM_FullService as Service,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] ,BM_FullServiceID from   (SELECT pm.current_Health,  (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN   JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold') and j.BM_FullService = ji.BM_FullService and  p.current_Health is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition' WHEN ji.BM_FullService = 2 THEN 'End to End' END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold') group by pm.current_Health,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                }
                if (sStatus == "DateRange")
                {
                    strQueryFinal = "select BM_FullService as Service,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] ,BM_FullServiceID from   (SELECT pm.current_Health,  (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN   JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) and j.BM_FullService = ji.BM_FullService and  p.current_Health is not null) as projects,CASE WHEN ji.BM_FullService = 0 THEN 'Full Service'  WHEN ji.BM_FullService = 1 THEN 'Composition' WHEN ji.BM_FullService = 2 THEN 'End to End' END as BM_FullService,ji.BM_FullService as BM_FullServiceID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Health,ji.BM_FullService) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by BM_FullService,projects,BM_FullServiceID";
                }
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[0].ToString(), "Service","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[1].ToString(), "Service","DashBoard",sStatus,"",""):a[1].ToString(),
                                     (a[2].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[2].ToString(), "Service","DashBoard",sStatus,"On Track",""):a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[3].ToString(), "Service","DashBoard",sStatus,"Watch Listed",""):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[4].ToString(), "Service","DashBoard",sStatus,"In Escalation",""):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[5].ToString(), "Service","DashBoard",sStatus,"N/A",""):a[5].ToString(),
                                     ""
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetDivisionListbyHealth(string sStatus, string sStartDate, string sEndDate)
        {
            //Project Health by Division
            try
            {
                string strQueryFinal = "select JBM_Location as Division,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],DivId from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "'  and j.JBM_Location = ji.JBM_Location and  p.current_Health is not null) as projects,CASE WHEN ji.JBM_Location = 0 THEN 'N/A'WHEN ji.JBM_Location = 1 THEN 'HSS'  WHEN ji.JBM_Location = 2 THEN 'STEM' END as JBM_Location,count(1) as statuscount,JBM_Location as DivId FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Health,ji.JBM_Location) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by JBM_Location,projects,DivId";
                if (sStatus == "In Progress")
                    strQueryFinal = "select JBM_Location as Division,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],DivId from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold') and j.JBM_Location = ji.JBM_Location and  p.current_Health is not null) as projects,CASE WHEN ji.JBM_Location = 0 THEN 'N/A'WHEN ji.JBM_Location = 1 THEN 'HSS'  WHEN ji.JBM_Location = 2 THEN 'STEM' END as JBM_Location,count(1) as statuscount,JBM_Location as DivId FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold') group by pm.current_Health,ji.JBM_Location) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by JBM_Location,projects,DivId";
                if (sStatus == "DateRange")
                    strQueryFinal = "select JBM_Location as Division,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],DivId from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) and j.JBM_Location = ji.JBM_Location and  p.current_Health is not null) as projects,CASE WHEN ji.JBM_Location = 0 THEN 'N/A'WHEN ji.JBM_Location = 1 THEN 'HSS'  WHEN ji.JBM_Location = 2 THEN 'STEM' END as JBM_Location,count(1) as statuscount,JBM_Location as DivId FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Health,ji.JBM_Location) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by JBM_Location,projects,DivId";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[0].ToString(), "Division","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[1].ToString(), "Division","DashBoard",sStatus,"",""):a[1].ToString(),
                                     (a[2].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[2].ToString(), "Division","DashBoard",sStatus,"On Track",""):a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[3].ToString(), "Division","DashBoard",sStatus,"Watch Listed",""):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[4].ToString(), "Division","DashBoard",sStatus,"In Escalation",""):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[5].ToString(), "Division","DashBoard",sStatus,"N/A",""):a[5].ToString(),
                                     ""
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
       [SessionExpire]
        public ActionResult GetSiteListbyHealth(string sStatus, string sStartDate, string sEndDate)
        {
            //Project Health by Site
            try
            {
                string strQueryFinal = "select Cenveo_Facility as Site,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '"+ Session["CustID"].ToString() + "'  and ISNULL(p.Cenveo_Facility, 'NULL') = ISNULL(pm.Cenveo_Facility, 'NULL') and  p.current_Health is not null) as projects,ISNULL(Cenveo_Facility, 'NULL') as Cenveo_Facility,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Health,pm.Cenveo_Facility) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by Cenveo_Facility,projects";
                if (sStatus == "In Progress")
                    strQueryFinal = "select Cenveo_Facility as Site,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold') and ISNULL(p.Cenveo_Facility, 'NULL') = ISNULL(pm.Cenveo_Facility, 'NULL') and  p.current_Health is not null) as projects,ISNULL(Cenveo_Facility, 'NULL') as Cenveo_Facility,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold') group by pm.current_Health,pm.Cenveo_Facility) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by Cenveo_Facility,projects";
                if (sStatus == "DateRange")
                    strQueryFinal = "select Cenveo_Facility as Site,projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) and ISNULL(p.Cenveo_Facility, 'NULL') = ISNULL(pm.Cenveo_Facility, 'NULL') and  p.current_Health is not null) as projects,ISNULL(Cenveo_Facility, 'NULL') as Cenveo_Facility,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Health,pm.Cenveo_Facility) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by Cenveo_Facility,projects";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[0].ToString(), "Site","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[1].ToString(), "Site","DashBoard",sStatus,"",""):a[1].ToString(),
                                     (a[2].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[2].ToString(), "Site","DashBoard",sStatus,"On Track",""):a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[3].ToString(), "Site","DashBoard",sStatus,"Watch Listed",""):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[4].ToString(), "Site","DashBoard",sStatus,"In Escalation",""):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[5].ToString(), "Site","DashBoard",sStatus,"N/A",""):a[5].ToString(),
                                     ""
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetPMListbyHealth(string sStatus, string sStartDate, string sEndDate)
        {
            //Project Health by Project Manager
            try
            {
                string strQueryFinal = "select ProjectCoordInd as [Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],PMID from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "'  and p.ProjectCoordInd = pm.ProjectCoordInd and  p.current_Health is not null) as projects,(SELECT EmpName FROM  JBM_EmployeeMaster where EmpLogin=pm.ProjectCoordInd) as ProjectCoordInd,ProjectCoordInd as PMID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Health,pm.ProjectCoordInd) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by ProjectCoordInd,projects,PMID";
                if (sStatus == "In Progress")
                    strQueryFinal = "select ProjectCoordInd as [Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],PMID from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold') and p.ProjectCoordInd = pm.ProjectCoordInd and  p.current_Health is not null) as projects,(SELECT EmpName FROM  JBM_EmployeeMaster where EmpLogin=pm.ProjectCoordInd) as ProjectCoordInd,ProjectCoordInd as PMID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold') group by pm.current_Health,pm.ProjectCoordInd) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by ProjectCoordInd,projects,PMID";
                if (sStatus == "DateRange")
                    strQueryFinal = "select ProjectCoordInd as [Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A],PMID from (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) and p.ProjectCoordInd = pm.ProjectCoordInd and  p.current_Health is not null) as projects,(SELECT EmpName FROM  JBM_EmployeeMaster where EmpLogin=pm.ProjectCoordInd) as ProjectCoordInd,ProjectCoordInd as PMID,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Health,pm.ProjectCoordInd) as a PIVOT(max(statuscount) FOR current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by ProjectCoordInd,projects,PMID";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[0].ToString(), "PM","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[1].ToString(), "PM","DashBoard",sStatus,"",""):a[1].ToString(),
                                     (a[2].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[2].ToString(), "PM","DashBoard",sStatus,"On Track",""):a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[3].ToString(), "PM","DashBoard",sStatus,"Watch Listed",""):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[4].ToString(), "PM","DashBoard",sStatus,"In Escalation",""):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[6].ToString(),Session["CustID"].ToString() , a[5].ToString(), "PM","DashBoard",sStatus,"N/A",""):a[5].ToString(),
                                     ""
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetCPMListbyHealth(string sStatus, string sStartDate, string sEndDate)
        {
            //Project Health by Customer Project Manager
            try
            {
                string strQueryFinal = "select ProjectManagerUS as [Customer Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from  (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '"+ Session["CustID"].ToString() + "'  and p.ProjectManagerUS = pm.ProjectManagerUS and  p.current_Health is not null) as projects,ProjectManagerUS,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by pm.current_Health,pm.ProjectManagerUS) as a PIVOT(max(statuscount) FOR  Current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by projects,ProjectManagerUS";
                if (sStatus == "In Progress")
                    strQueryFinal = "select ProjectManagerUS as [Customer Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from  (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and p.Current_Status in ('In Progress','On Hold') and p.ProjectManagerUS = pm.ProjectManagerUS and  p.current_Health is not null) as projects,ProjectManagerUS,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and pm.Current_Status in ('In Progress','On Hold') group by pm.current_Health,pm.ProjectManagerUS) as a PIVOT(max(statuscount) FOR  Current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by projects,ProjectManagerUS";
                if (sStatus == "DateRange")
                    strQueryFinal = "select ProjectManagerUS as [Customer Project Manager],projects,sum(ISNULL([On Track], 0)) as [On Track],sum(ISNULL([Watch Listed], 0)) as [Watch Listed],sum(ISNULL([In Escalation], 0)) as [In Escalation], sum(ISNULL([N/A], 0)) as [N/A] from  (SELECT pm.current_Health, (SELECT count(p.JBM_AutoID) as projects FROM  BK_ProjectManagement p JOIN  JBM_Info j ON p.JBM_AutoID = j.JBM_AutoID where j.CustID = '" + Session["CustID"].ToString() + "' and j.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) and p.ProjectManagerUS = pm.ProjectManagerUS and  p.current_Health is not null) as projects,ProjectManagerUS,count(1) as statuscount FROM BK_ProjectManagement pm JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID where ji.CustID = '" + Session["CustID"].ToString() + "' and pm.current_Health is not null and ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and ji.JBM_PrinterDate between CONVERT(DATETIME, '" + sStartDate + "', 101)  and CONVERT(DATETIME, '" + sEndDate + "',101) group by pm.current_Health,pm.ProjectManagerUS) as a PIVOT(max(statuscount) FOR  Current_Health IN([On Track],[Watch Listed],[In Escalation],[N/A])) AS dt group by projects,ProjectManagerUS";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[0].ToString(), "CPM","DashBoard",sStatus,"",""),
                                     (a[1].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[1].ToString(), "CPM","DashBoard",sStatus,"",""):a[1].ToString(),
                                     (a[2].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[2].ToString(), "CPM","DashBoard",sStatus,"On Track",""):a[2].ToString(),
                                     (a[3].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[3].ToString(), "CPM","DashBoard",sStatus,"Watch Listed",""):a[3].ToString(),
                                     (a[4].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[4].ToString(), "CPM","DashBoard",sStatus,"In Escalation",""):a[4].ToString(),
                                     (a[5].ToString()!="0")?CreateHREFLinkforDashboard(Session["sCustSN"].ToString(),a[0].ToString(),Session["CustID"].ToString() , a[5].ToString(), "CPM","DashBoard",sStatus,"N/A",""):a[5].ToString(),
                                     ""
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetRemarksList(string sSubdivision, string sProduction, string sFacilities, string sPMProject)
        {
            try
            {
                Session["sPMProject"] = sPMProject;
                string strQueryFinal = "SELECT ji.JBM_AutoID, si.Instruction,si.InstDate,ji.JBM_ID,ji.DocketNo,pm.Cenveo_Facility,pm.DesignDesc,(select EmpName from JBM_EmployeeMaster where Emplogin=pm.ProductionLead) as ProductionLead,pm.BusinessUnit as [Subdivision],jc.CustSN,ji.JBM_Intrnl,pm.Current_Health,(select empname from JBM_employeemaster where emplogin = pm.ProjectCoordInd) as PM,Convert(varchar(11),Last_Review,106) as [Last_Review] FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID left JOIN BK_SplInstructions si ON pm.JBM_AutoID = si.AutoArtID and si.InstDate = (select max(InstDate) from BK_SplInstructions bki where bki.AutoArtID=ji.JBM_AutoID) where ji.jbm_disabled='0' and pm.current_Status is not null and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                //string strResult = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString() );

                //if (strResult != "-1")
                //{
                //    strQueryFinal += " and ji.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "')";
                //}

                string sCustSN = Session["sesCustSN"].ToString();
                string sCPM = Session["sesCPM"].ToString();
                //string sPMProject = Session["sesPMProject"].ToString();
                string sServices = Session["sesServices"].ToString();
                string sStatus = Session["sesStatus"].ToString();
                string sHealth = Session["sesHealth"].ToString();

                if (sCustSN != "AllCustomers")
                {
                    strQueryFinal += " and jc.CustSN in ('" + sCustSN.Trim() + "')";
                }

                if (sServices != "AllServices")
                {
                    strQueryFinal += " and ji.BM_FullService = '" + sServices.Trim() + "'";
                }

                if (sStatus != "AllStatus")
                {
                    if (sStatus != "Live Projects")
                    {
                        strQueryFinal += " and pm.Current_Status = '" + sStatus.Trim() + "'";
                    }
                    else
                    {
                        strQueryFinal += " and pm.Current_Status in('In Progress','On Hold','Yet to Start')";
                    }
                }

                if (sHealth != "AllHealth")
                {
                    strQueryFinal += " and pm.Current_Health = '" + sHealth.Trim() + "'";
                }

                if (sCPM != "AllCPMs")
                {
                    strQueryFinal += " and pm.ProjectManagerUS = '" + sCPM.Trim() + "'";
                }

                if (sPMProject == "MyProjects" && Session["CustomerSN"].ToString() == "")   //LPM/PM
                {
                    strQueryFinal += " and pm.ProjectCoordInd = '" + Session["EmpLogin"].ToString().Trim() + "'";
                }
                else if (sPMProject == "MyTeam") //PM/LPM
                {
                    strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' ";
                    //strQueryFinal += " and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "'  and  pm.ProjectCoordInd in (select emplogin from JBM_Employeemaster where usergroup=(select emplogin from JBM_Employeemaster where emplogin='" + Session["EmpLogin"].ToString().Trim() + "' and roleid='103') and roleid='104') ";
                }
                if (!Regex.IsMatch(sPMProject, "(MyProjects|MyTeam|AllProjects|AllPMs)", RegexOptions.IgnoreCase))
                {
                    strQueryFinal += "  and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                    //strQueryFinal += "and ji.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' and pm.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                }
                if (sSubdivision != "AllSubdivision")
                {
                    strQueryFinal += " and pm.BusinessUnit = '" + sSubdivision.Trim() + "'";
                }

                if (sProduction != "AllProduction")
                {
                    strQueryFinal += " and pm.ProductionLead = '" + sProduction.Trim() + "'";
                }

                if (sFacilities != "AllFacilities")
                {
                    strQueryFinal += " and pm.Cenveo_Facility = '" + sFacilities.Trim() + "'";
                }

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by JBM_ID asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateHREFLink(a[0].ToString(),"HrefLink", a[9].ToString(), a[10].ToString()),
                                     a[12].ToString(),
                                     dt_DateFrmtSort(a[13].ToString()) + "<input type='text' class='form-control datepicker' value='" + a[13].ToString() + "' data-jid='" + a[0].ToString() + "'/>",
                                     RemarkText(a[1].ToString()),
                                     a[6].ToString().Trim()=="[Select]"?"":a[6].ToString().Trim(),
                                     a[8].ToString().Trim()=="[Select]"?"":a[8].ToString().Trim(),
                                     a[4].ToString(),
                                     a[7].ToString(),
                                     a[5].ToString(),
                                     HighlightHealth(a[11].ToString(), "Remarks")
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public string updateLastReview(string JID, string LastReview)
        {
            try
            {
                return DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_ProjectManagement set Last_Review= Nullif('" + LastReview.Trim() + "','') where JBM_AutoID='" + JID + "'; Select top 1 Convert(varchar(11),Last_Review,106) as Last_Review from " + Session["sCustAcc"].ToString() + "_ProjectManagement where JBM_AutoID='" + JID + "'", Session["sConnSiteDB"].ToString());
            }
            catch(Exception ex)
            {
                return "";
            }
        }
        public string RemarkText(string Notes) {
            try
            {
                string strNotes = string.Empty;
                if (Notes.Length > 80)
                {
                    strNotes = "<div class='remarksdes'>" + Notes + "</div>";
                }
                else {
                    strNotes = Notes;
                }
                
                return strNotes;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult Schedule()
        {
            try
            {
                if (Session["CustomerSN"].ToString() != "")
                {
                    string sCustSN = Session["CustomerSN"].ToString();
                    DataSet ds1 = new DataSet();
                    string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and jc.CustSN='" + sCustSN + "' order by CustSN asc ";
                    ds1 = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        ViewBag.vCustID = ds1.Tables[0].Rows[0]["CustID"].ToString();
                        ViewBag.vCustSN = ds1.Tables[0].Rows[0]["CustSN"].ToString();
                        ViewBag.vCustName = ds1.Tables[0].Rows[0]["CustName"].ToString();
                    }
                }

                if (Session["SchedulePageLength"] == null)
                {
                    Session["SchedulePageLength"] = 10;
                } 

                ViewBag.PrjTabHead = "Projects";
                ViewBag.PageHead = "Project Tracking - Schedule";
                List<SelectListItem> Useritems = new List<SelectListItem>();

                Useritems.Add(new SelectListItem { Text = "Copy Editing", Value = "CE" });
                Useritems.Add(new SelectListItem { Text = "First Pages", Value = "FP" });
                Useritems.Add(new SelectListItem { Text = "Second Pages", Value = "2ndPg" });
                Useritems.Add(new SelectListItem { Text = "Final Pages", Value = "FinPag" });

                ViewBag.Userlist = Useritems;

                //Load Project Manager 
                List<SelectListItem> lstPM = new List<SelectListItem>();
                DataSet dsPM = new DataSet();
                dsPM = DBProc.GetResultasDataSet("select EmpName,emplogin from JBM_Employeemaster where usergroup = (select emplogin from JBM_Employeemaster where emplogin = '" + Session["EmpLogin"].ToString().Trim() + "' and roleid = '103') and roleid = '104'", Session["sConnSiteDB"].ToString());
                if (dsPM.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPM.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpName = dsPM.Tables[0].Rows[intCount]["EmpName"].ToString();
                        string stremplogin = dsPM.Tables[0].Rows[intCount]["emplogin"].ToString();

                        lstPM.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = stremplogin.ToString()
                        });
                    }

                }

                ViewBag.PMlist = lstPM;
                //Load Project Manager list Items
                DataSet ds = new DataSet();
                List<SelectListItem> itemsProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpAutoID,EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='104' or roleid='103' order by empname asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.PrjMgrList = itemsProjectManager;
                return View();

            }
            catch (Exception)
            {

                throw;
            }


        }
        [SessionExpire]
        public ActionResult GetScheduleList(string sPMProject)
        {
            try
            {
                Session["sPMProject"] = sPMProject;
                string sCustSN = Session["sesCustSN"].ToString();
                string sCPM = Session["sesCPM"].ToString();
                //string sPMProject = Session["sesPMProject"].ToString();
                string sServices = Session["sesServices"].ToString();
                string sStatus = Session["sesStatus"].ToString();
                string sHealth = Session["sesHealth"].ToString();

                DataSet ds = new DataSet();
                string sEmpLogin = Session["EmpLogin"].ToString().Trim();
                string strQueryFinal = @"select distinct parent.JBM_AutoID,parent.JBM_Intrnl,parent.CustSN,FORMAT(CAST(ce.PlannedStartDate as date), 'dd-MMM-yy')
as CE_PlannedStartDate,FORMAT(CAST(ce.PlannedEndDate as date), 'dd-MMM-yy') as CE_PlannedEndDate,FORMAT(CAST(ce.ReceivedDate as date), 'dd-MMM-yy') 
as CE_ReceivedDate,FORMAT(CAST(ce.DispatchDate as date), 'dd-MMM-yy') as CE_DispatchDate,FORMAT(CAST(ce.RevisedStartDate as date), 'dd-MMM-yy') as CE_RevisedStartDate,
FORMAT(CAST(ce.RevisedEndDate as date), 'dd-MMM-yy') as CE_RevisedEndDate,FORMAT(CAST(fp.PlannedStartDate as date), 'dd-MMM-yy') 
as FP_PlannedStartDate,FORMAT(CAST(fp.PlannedEndDate as date), 'dd-MMM-yy') as FP_PlannedEndDate,FORMAT(CAST(fp.ReceivedDate as date), 'dd-MMM-yy') 
as FP_ReceivedDate,FORMAT(CAST(fp.DispatchDate as date), 'dd-MMM-yy') as FP_DispatchDate,FORMAT(CAST(fp.RevisedStartDate as date), 'dd-MMM-yy') as FP_RevisedStartDate,
FORMAT(CAST(fp.RevisedEndDate as date), 'dd-MMM-yy') as FP_RevisedEndDate,FORMAT(CAST(FinPag.PlannedStartDate as date), 'dd-MMM-yy') 
as FinPag_PlannedStartDate,FORMAT(CAST(FinPag.PlannedEndDate as date), 'dd-MMM-yy') as FinPag_PlannedEndDate,FORMAT(CAST(FinPag.ReceivedDate as date), 'dd-MMM-yy') 
as FinPag_ReceivedDate,FORMAT(CAST(FinPag.DispatchDate as date), 'dd-MMM-yy') as FinPag_DispatchDate,FORMAT(CAST(FinPag.RevisedStartDate as date), 'dd-MMM-yy') as FinPag_RevisedStartDate,
FORMAT(CAST(FinPag.RevisedEndDate as date), 'dd-MMM-yy') as FinPag_RevisedEndDate,FORMAT(CAST(secondPg.PlannedStartDate as date), 'dd-MMM-yy') 
as secondPg_PlannedStartDate,FORMAT(CAST(secondPg.PlannedEndDate as date), 'dd-MMM-yy') as secondPg_PlannedEndDate,FORMAT(CAST(secondPg.ReceivedDate as date), 'dd-MMM-yy') 
as secondPg_ReceivedDate,FORMAT(CAST(secondPg.DispatchDate as date), 'dd-MMM-yy') as secondPg_DispatchDate,FORMAT(CAST(secondPg.RevisedStartDate as date), 'dd-MMM-yy') as secondPg_RevisedStartDate,
FORMAT(CAST(secondPg.RevisedEndDate as date), 'dd-MMM-yy') as secondPg_RevisedEndDate,FORMAT(CAST(parent.JBM_PrinterDate as date), 'dd-MMM-yy')as JBM_PrinterDate,
parent.Current_Health,(select empname from JBM_employeemaster where emplogin=parent.ProjectCoordInd) as PM
 from 
(SELECT min(SI.DueDate) as PlannedStartDate, max(SI.DueDate) as PlannedEndDate,min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate,min(SI.RevisedDate) as RevisedStartDate, max(SI.RevisedDate) as RevisedEndDate,pm.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl ,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS ,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName from BK_ProjectManagement pm left join BK_ChapterInfo CI
ON pm.JBM_AutoID = CI.JBM_AutoID left JOIN BK_Stageinfo SI
ON CI.AutoArtID = SI.AutoArtID left JOIN JBM_Info ji 
ON pm.JBM_AutoID = ji.JBM_AutoID left JOIN JBM_CustomerMaster jc
ON ji.CustID = jc.CustID group by pm.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName) parent
left join
(SELECT min(SI.CeDueDate) as PlannedStartDate, max(SI.CeDueDate) as PlannedEndDate,min(SI.CeDispDate) as ReceivedDate, max(SI.CeDispDate) as DispatchDate,min(SI.CERevisedDate) as RevisedStartDate, max(SI.CERevisedDate) as RevisedEndDate,pm.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl ,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS ,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName from BK_ProjectManagement pm left join BK_ChapterInfo CI
ON pm.JBM_AutoID = CI.JBM_AutoID left JOIN BK_Stageinfo SI
ON CI.AutoArtID = SI.AutoArtID left JOIN JBM_Info ji 
ON pm.JBM_AutoID = ji.JBM_AutoID left JOIN JBM_CustomerMaster jc
ON ji.CustID = jc.CustID where SI.RevFinStage = 'FP' group by pm.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName ) ce on ce.JBM_AutoID=parent.JBM_AutoID and ce.Short_Stage='FP'
left join
(SELECT min(SI.DueDate) as PlannedStartDate, max(SI.DueDate) as PlannedEndDate,min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate,min(SI.RevisedDate) as RevisedStartDate, max(SI.RevisedDate) as RevisedEndDate,pm.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl ,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS ,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName from BK_ProjectManagement pm left join BK_ChapterInfo CI
ON pm.JBM_AutoID = CI.JBM_AutoID left JOIN BK_Stageinfo SI
ON CI.AutoArtID = SI.AutoArtID left JOIN JBM_Info ji 
ON pm.JBM_AutoID = ji.JBM_AutoID left JOIN JBM_CustomerMaster jc
ON ji.CustID = jc.CustID group by pm.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName) fp on fp.JBM_AutoID=parent.JBM_AutoID and fp.Short_Stage='FP'
left join
(SELECT min(SI.DueDate) as PlannedStartDate, max(SI.DueDate) as PlannedEndDate,min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate,min(SI.RevisedDate) as RevisedStartDate, max(SI.RevisedDate) as RevisedEndDate,pm.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl ,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS ,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName from BK_ProjectManagement pm left join BK_ChapterInfo CI
ON pm.JBM_AutoID = CI.JBM_AutoID left JOIN BK_Stageinfo SI
ON CI.AutoArtID = SI.AutoArtID left JOIN JBM_Info ji 
ON pm.JBM_AutoID = ji.JBM_AutoID left JOIN JBM_CustomerMaster jc
ON ji.CustID = jc.CustID group by pm.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName  ) FinPag on FinPag.JBM_AutoID=parent.JBM_AutoID and FinPag.Short_Stage='FinPag'
 left join
(SELECT min(SI.DueDate) as PlannedStartDate, max(SI.DueDate) as PlannedEndDate,min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate,min(SI.RevisedDate) as RevisedStartDate, max(SI.RevisedDate) as RevisedEndDate,pm.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl ,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS ,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName from BK_ProjectManagement pm left join BK_ChapterInfo CI
ON pm.JBM_AutoID = CI.JBM_AutoID left JOIN BK_Stageinfo SI
ON CI.AutoArtID = SI.AutoArtID left JOIN JBM_Info ji 
ON pm.JBM_AutoID = ji.JBM_AutoID left JOIN JBM_CustomerMaster jc
ON ji.CustID = jc.CustID group by pm.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl,ji.JBM_PrinterDate,pm.ProjectCoordInd,pm.ProjectManagerUS,
pm.Current_Status,pm.Current_Health,ji.BM_FullService,ji.jbm_disabled,ji.KGLAccMgrName  ) secondPg on secondPg.JBM_AutoID=parent.JBM_AutoID 
and secondPg.Short_Stage='2ndPg' where parent.jbm_disabled = '0' and parent.current_Status is not null and parent.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";

                if (sCustSN != "AllCustomers")
                {
                    strQueryFinal += " and parent.CustSN in ('" + sCustSN.Trim() + "')";
                }

                if (sServices != "AllServices")
                {
                    strQueryFinal += " and parent.BM_FullService = '" + sServices.Trim() + "'";
                }

                if (sStatus != "AllStatus")
                {
                    if(sStatus!= "Live Projects")
                    {
                        strQueryFinal += " and parent.Current_Status = '" + sStatus.Trim() + "'";
                    }
                    else
                    {
                        strQueryFinal += " and parent.Current_Status in('In Progress','On Hold','Yet to Start')";
                    }
                }

                if (sHealth != "AllHealth")
                {
                    strQueryFinal += " and parent.Current_Health = '" + sHealth.Trim() + "'";
                }

                if (sCPM != "AllCPMs")
                {
                    strQueryFinal += " and parent.ProjectManagerUS = '" + sCPM.Trim() + "'";
                }

                if (sPMProject == "MyProjects" && Session["CustomerSN"].ToString() == "")   //LPM/PM
                {
                    strQueryFinal += " and parent.ProjectCoordInd = '" + sEmpLogin.Trim() + "'";
                }
                else if (sPMProject == "MyTeam") //PM/LPM
                {
                    strQueryFinal += " and parent.KGLAccMgrName = '" + sEmpLogin.Trim() + "' ";
                    //strQueryFinal += " and parent.KGLAccMgrName = '" + sEmpLogin.Trim() + "'  and  parent.ProjectCoordInd in (select emplogin from JBM_Employeemaster where usergroup=(select emplogin from JBM_Employeemaster where emplogin='" + sEmpLogin.Trim() + "' and roleid='103') and roleid='104') ";
                }
                if (!Regex.IsMatch(sPMProject, "(MyProjects|MyTeam|AllProjects|AllPMs)", RegexOptions.IgnoreCase))
                {
                    strQueryFinal += "  and parent.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                    //strQueryFinal += "and parent.KGLAccMgrName = '" + Session["EmpLogin"].ToString().Trim() + "' and parent.ProjectCoordInd = '" + sPMProject.Trim() + "'";
                }
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {CreateHREFLink(a[0].ToString(),"HrefLink", a[2].ToString(), a[1].ToString()),
                                     a[2].ToString(),a[29].ToString(),
                                     dt_DateFrmtSort(a[3].ToString()) + a[3].ToString(),   //CE_PlannedStartDate     // Copyediting
                                     dt_DateFrmtSort(a[4].ToString()) + a[4].ToString(),   //CE_PlannedEndDate
                                     dt_DateFrmtSort(a[7].ToString()) + a[7].ToString(),  //CE_RevisedStartDate
                                     dt_DateFrmtSort(a[8].ToString()) + a[8].ToString(),    //CE_RevisedEndDate
                                     dt_DateFrmtSort(a[5].ToString()) + a[5].ToString(),   //CE_ReceivedDate
                                     dt_DateFrmtSort(a[6].ToString()) + a[6].ToString(),  //CE_DispatchDate                                    

                                     dt_DateFrmtSort(a[9].ToString()) + a[9].ToString(),   ///FP_PlannedStartDate   // FP Pages
                                     dt_DateFrmtSort(a[10].ToString()) + a[10].ToString(),   //FP_PlannedEndDate
                                     dt_DateFrmtSort(a[13].ToString()) + a[13].ToString(),   //FP_RevisedStartDate
                                     dt_DateFrmtSort(a[14].ToString()) + a[14].ToString(),   //FP_RevisedEndDate
                                     dt_DateFrmtSort(a[11].ToString()) + a[11].ToString(),   //FP_ReceivedDate
                                     dt_DateFrmtSort(a[12].ToString()) + a[12].ToString(),  //FP_DispatchDate                                    
                                     
                                     dt_DateFrmtSort(a[21].ToString()) + a[21].ToString(),  //secondPg_PlannedStartDate     // Second Pages
                                     dt_DateFrmtSort(a[22].ToString()) + a[22].ToString(),  //secondPg_PlannedEndDate                                     
                                     dt_DateFrmtSort(a[25].ToString()) + a[25].ToString(),   //secondPg_RevisedStartDate
                                     dt_DateFrmtSort(a[26].ToString()) + a[26].ToString(),   //secondPg_RevisedEndDate
                                     dt_DateFrmtSort(a[23].ToString()) + a[23].ToString(),  //secondPg_ReceivedDate
                                     dt_DateFrmtSort(a[24].ToString()) + a[24].ToString(),  //secondPg_DispatchDate

                                     dt_DateFrmtSort(a[15].ToString()) + a[15].ToString(),  //FinPag_PlannedStartDate   //Final pages
                                     dt_DateFrmtSort(a[16].ToString()) + a[16].ToString(),  //FinPag_PlannedEndDate                                     
                                     dt_DateFrmtSort(a[19].ToString()) + a[19].ToString(),   //FinPag_RevisedStartDate
                                     dt_DateFrmtSort(a[20].ToString()) + a[20].ToString(),   //FinPag_RevisedEndDate
                                     dt_DateFrmtSort(a[17].ToString()) + a[17].ToString(),  //FinPag_ReceivedDate
                                     dt_DateFrmtSort(a[18].ToString()) + a[18].ToString(),  //FinPag_DispatchDate

                                     dt_DateFrmtSort(a[27].ToString()) + a[27].ToString(),  //JBM_PrinterDate
                                     HighlightHealth(a[28].ToString().Trim(), "Schedule")   //Current_Health
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult DashBoard(string CustAcc, string sCustSN, string EmpID, string SiteID)
        {
            if (Session["CustomerSN"].ToString() != "")
            {
                sCustSN = Session["CustomerSN"].ToString();
            }
             
           // Session["DashboardLink"] = "?CustAcc="+ CustAcc + "&amp;sCustSN=" + sCustSN + "&EmpID=" + EmpID + "&SiteID=" + SiteID + "";
            Session["sCustSN"] = sCustSN;
            DataSet ds = new DataSet();

            if (CustAcc != null)
            {
                string strUrl = Request.Url.AbsoluteUri.ToString();

                Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";
                Session["strHomeURL"] = strUrl;
                Session["EmpIdLogin"] = EmpID;
                Session["sCustAcc"] = CustAcc;
                Session["sSiteID"] = SiteID;

                clsCollec.getSiteDBConnection(SiteID, CustAcc);
                if (Session["sConnSiteDB"].ToString() == "")
                {
                    Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                }


                ds = DBProc.GetResultasDataSet("Select EmpLogin,EmpAutoId,EmpName,EmpMailId,DeptCode,RoleId, DeptAccess,CustAccess,TeamID,(Select b.DeptName from JBM_DepartmentMaster b  where e.DeptCode = b.DeptCode) as DeptName from JBM_EmployeeMaster e Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
                    Session["EmpLogin"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
                    Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();
                    Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                    Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                    Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();

                    GlobalVariables.strEmpName = ds.Tables[0].Rows[0]["EmpName"].ToString();

                }
            }

            ds = new DataSet();
            string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and jc.CustSN='"+sCustSN+"' order by CustSN asc ";
            ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                Session["CustID"] = ds.Tables[0].Rows[0]["CustID"].ToString();
                Session["CustName"] = ds.Tables[0].Rows[0]["CustName"].ToString();
                ViewBag.vCustID = ds.Tables[0].Rows[0]["CustID"].ToString();
                ViewBag.vCustSN = ds.Tables[0].Rows[0]["CustSN"].ToString();
                ViewBag.vCustName = ds.Tables[0].Rows[0]["CustName"].ToString();

            }

           

            ViewBag.PrjTabHead = "Project Tracking";
            ViewBag.PageHead = "Project Tracking - Dashboard";
            return View();

        }
        [SessionExpire]
        public ActionResult Customer()
        {
            try
            {
                List<SelectListItem> lstCustomer = new List<SelectListItem>();
                DataSet ds = new DataSet();
                string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName,jc.CustID from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' order by CustSN asc ";
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        lstCustomer.Add(new SelectListItem
                        {
                            Text = ds.Tables[0].Rows[intCount]["CustName"].ToString(),
                            Value = ds.Tables[0].Rows[intCount]["CustSN"].ToString()
                        });
                    }

                }

                Session["CustomerSN"] = "";
                ViewBag.Customerlist = lstCustomer;

                ViewBag.PrjTabHead = "Project Tracking";
                ViewBag.PageHead = "Project Tracking - Customer";
                return View();
            }
            catch (Exception)
            {
                return View();
            }

            

        }

        [SessionExpire]
        public ActionResult GetCustomerList()
        {
            try
            {
                string strQueryFinal = "Select  DISTINCT jc.CustSN, jc.CustName from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' ";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by CustSN asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString(),
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
        public ActionResult AddProjects()
        {
            try
            {
                //Load Customer list Items
                List<SelectListItem> lstCustomer = new List<SelectListItem>();
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select DISTINCT jc.CustID,jc.CustSN from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' order by CustSN asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        string strCustSN = ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                        string strCustID = ds.Tables[0].Rows[intCount]["CustID"].ToString();
                        lstCustomer.Add(new SelectListItem
                        {
                            Text = strCustSN.ToString(),
                            Value = strCustID.ToString()
                        });
                    }

                }

                ViewBag.Customerlist = lstCustomer;

                //Load Lead Project Manager list Items
                List<SelectListItem> itemsLeadProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='103' order by EmpName asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsLeadProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.LeadPrjMgrList = itemsLeadProjectManager;

                //Load Project Manager list Items
                List<SelectListItem> itemsProjectManager = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='104' or roleid='103' order by EmpName asc", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsProjectManager.Add(new SelectListItem
                        {
                            Text = myRow["EmpName"].ToString(),
                            Value = myRow["EmpLogin"].ToString()
                        });
                    }
                }
                ViewBag.PrjMgrList = itemsProjectManager;


                //Load KGL Facility list Items
                List<SelectListItem> itemsKGLFacility = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select SiteId,SiteName from [dbo].[JBM_SiteMaster] where SiteName not in ('Cadmus')", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        itemsKGLFacility.Add(new SelectListItem
                        {
                            Text = myRow["SiteName"].ToString(),
                            Value = myRow["SiteName"].ToString()
                        });
                    }
                    itemsKGLFacility.Add(new SelectListItem { Text = "Richmond, VA", Value = "Richmond, VA" });
                    //itemsKGLFacility.Add(new SelectListItem { Text = "Columbia, MD", Value = "Columbia, MD" }); // Ben requested 
                    //itemsKGLFacility.Add(new SelectListItem { Text = "Fort Washington, PA", Value = "Fort Washington, PA" });
                    itemsKGLFacility.Add(new SelectListItem { Text = "London, UK", Value = "London, UK" });
                }
                ViewBag.KGLFacilityList = itemsKGLFacility;
            }
            catch { }
            ViewBag.PrjTabHead = "Project Tracking";
            ViewBag.PageHead = "Project Tracking - Add Project";

            return View();

        }

        [SessionExpire]
        public ActionResult GetCustomerPMList(string sCustomer)
        {

            try
            {
                if (sCustomer != "" || sCustomer != null)
                {
                    string strQueryFinal = "SELECT distinct pm.ProjectManagerUS as CPM FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and ji.JBM_AutoID like 'BK%' and jc.custsn = '" + sCustomer + "' and pm.ProjectManagerUS is not null";

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
        public ActionResult AddNewProject(string sCustID, string sPMID, string sLPMID, string sSID, string sKGLID, string sCPMID, string sPID, string sJDocket, string sTitle, string sEdtion, string sAuthor, string sISBN,string sPTID)
        {
            try
            {
                string strJBMAutoMaxID = "";
                string strAutoSeq = "";
                string strAutoSeq1 = "";
                string JBM_IntrnlID = "";
                string JBM_Intrnl = "";

                string strAutoIdFilter = Session["sCustAcc"].ToString();

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl("Select max(convert(int,substring(JBM_AutoID,3,9))) as JBM_AutoID from JBM_Info where JBM_AutoID like '" + strAutoIdFilter + "%'", Session["sConnSiteDB"].ToString());

                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["JBM_AutoID"].ToString() != null || dt.Rows[0]["JBM_AutoID"].ToString().Trim() != "")
                    {
                        // strJBMAutoMaxID = (Convert.ToInt32(dt.Rows[0]["JBM_AutoID"].ToString().Substring(2)) + 1).ToString();
                        strJBMAutoMaxID = dt.Rows[0]["JBM_AutoID"].ToString();
                        strJBMAutoMaxID = (Convert.ToInt32(strJBMAutoMaxID) + 1).ToString();
                    }
                    switch (strJBMAutoMaxID.Length)
                    {
                        case 1:
                            strJBMAutoMaxID = strAutoIdFilter + "00" + strJBMAutoMaxID;
                            break;
                        case 2:
                            strJBMAutoMaxID = strAutoIdFilter + "0" + strJBMAutoMaxID;
                            break;
                        default:
                            strJBMAutoMaxID = strAutoIdFilter + strJBMAutoMaxID;
                            break;
                    }
                }
                else
                {
                    strJBMAutoMaxID = strAutoIdFilter + "001";
                }

                dt = DBProc.GetResultasDataTbl("Select max(AutoSeq) as AutoSeq from JBM_Info where JBM_AutoID like '" + strAutoIdFilter + "%' and JBM_IntrnlID like '" + (DateTime.Now.Year).ToString().Substring(2) + "%'", Session["sConnSiteDB"].ToString());

                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["AutoSeq"].ToString() != null || dt.Rows[0]["AutoSeq"].ToString().Trim() != "")
                    {
                        strAutoSeq = dt.Rows[0]["AutoSeq"].ToString();
						if (strAutoSeq == "" || strAutoSeq == null)
                            {
                            strAutoSeq = "0";
                        }
                        strAutoSeq1 = (Convert.ToInt32(strAutoSeq) + 1).ToString();
                        strAutoSeq = (Convert.ToInt32(strAutoSeq) + 1).ToString();
                    }
                }
                else
                {
                    strAutoSeq = "1";
                    strAutoSeq1 = "1";
                }

                JBM_IntrnlID = (DateTime.Now.Year).ToString().Substring(2) + strAutoSeq.PadLeft(4, '0');
                JBM_Intrnl = sPID.ToUpper().ToString().Trim() + "_" + sISBN.Replace("-", "").Replace(" ", "").Trim() + "_" + JBM_IntrnlID;

                //To validate project is already exists or not
                dt = DBProc.GetResultasDataTbl("Select JBM_Intrnl from JBM_Info where JBM_Intrnl like '" + sPID.Trim() + "_%' and CustID='" + sCustID + "'", Session["sConnSiteDB"].ToString());

                if (dt.Rows.Count > 0)
                {
                    return Json(new { dataComp = "This " + sPID.Trim() + " project is already exists." }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //SqlConnection con = new SqlConnection();
                    //con = DBProc.getConnection(Session["sConnSiteDB"].ToString() );
                    //con.Open();
                    DataSet ds = new DataSet();
                    string sEmpLogin = Session["EmpLogin"].ToString().Trim();

                    string strinsertJBM_Info = "INSERT INTO JBM_Info (AutoSeq,CustID,Docketno,JBM_AutoID,JBM_IntrnlID,JBM_ID,JBM_Intrnl,JBM_TeamID,JBM_SubTeam,Sample_wf,Fp_wf,Rev_wf,Fin_wf,Title,BM_Author,BM_ISBNnumber13,BM_ISBN10number,BM_FullService,JBM_Platform,JBM_PrinterDate,CopyrightOwner,JBM_Trimsize,BM_DesiredPgCount,KGLAccMgrName,JBM_Location,IndexerName,MssPages,JBM_disabled) values  ('" + strAutoSeq + "','" + sCustID + "','" + sJDocket + "','" + strJBMAutoMaxID + "','" + JBM_IntrnlID + "','" + JBM_Intrnl + "','" + JBM_Intrnl + "','3','12','W4','W192','W4','W4','" + sTitle.Replace("'", "''").ToString() + "','" + sAuthor.Replace("'", "''").ToString() + "','" + sISBN + "','','" + sSID + "','1',NULL,'','','0','" + sLPMID + "','0','','','0')";
                    string strResult2 = DBProc.GetResultasString(strinsertJBM_Info, Session["sConnSiteDB"].ToString());

                    string strinsertBK_ProjectMgnt = "INSERT INTO BK_ProjectManagement (JBM_AutoID,Edition,ProjectManagerUS,ProjectCoordInd,Cenveo_Facility,Current_Status,StageID) values('" + strJBMAutoMaxID + "','" + sEdtion.Replace("'", "''").ToString() + "','" + sCPMID + "','" + sPMID + "','" + sKGLID + "','Yet to Start','" + sPTID + "')";
                    string strResult1 = DBProc.GetResultasString(strinsertBK_ProjectMgnt, Session["sConnSiteDB"].ToString());


                    //SqlCommand cmdstrinsertBK_ProjectMgnt = new SqlCommand(strinsertBK_ProjectMgnt, con);
                    //cmdstrinsertBK_ProjectMgnt.ExecuteNonQuery();
                    //SqlCommand cmdstrinsertJBM_Info = new SqlCommand(strinsertJBM_Info, con);
                    //cmdstrinsertJBM_Info.ExecuteNonQuery();
                    //con.Close();
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }

               
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult AddHolidays()
        {
            ViewBag.PrjTabHead = "Projects";
            ViewBag.PageHead = "Project Tracking - Holidays List";
            return View();
        }
        [SessionExpire]
        public ActionResult GetHolidaysList(string sYear, string sFacilities)
        {
            try
            {
                string strQueryFinal = "Select ID,FORMAT (CAST(HolidayDate as date), 'dd-MMM-yyyy') as HolidayDate, Description,Years,Location from [dbo].[tbl_HolidayList] WHERE Years !=''";

                if (sYear != "AllYears")
                {
                    strQueryFinal += " and Years = '" + sYear.Trim() + "'";
                }

                if (sFacilities != "AllFacilities")
                {
                    strQueryFinal += " and Location = '" + sFacilities.Trim() + "'";
                }

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  Order by HolidayDate asc", Session["sConnSiteDB"].ToString());


                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     dt_DateFrmtSort(a[1].ToString()) + a[1].ToString(), 
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     "<button type='button' id='btnDelete" + a[0].ToString() + "' onClick=\"funcDelete('" +  a[0].ToString() + "')\"  class='btn btn-light' name='delete' value='Delete'><span class='fas fa-trash fa-1x text-red'></span></button>"
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult AddHolidaysList(string sDate, string sDesc, string sYear, string sFacilities)
        {
            try
            {

                sDesc = string.IsNullOrEmpty(sDesc) ? "" : sDesc.ToString().Replace("'", "''").ToString();
                string strResult = DBProc.GetResultasString("INSERT INTO [tbl_HolidayList] (HolidayDate,Description,Years,location) VALUES ('" + sDate + "','" + sDesc + "','" + sYear + "','" + sFacilities   + "')", Session["sConnSiteDB"].ToString());
                if (strResult == "1")
                {
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else
                { return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet); }
                
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult DeleteHolidayList(string sID)
        {
            try
            {
                string strResult = DBProc.GetResultasString("DELETE FROM [tbl_HolidayList] WHERE ID='" + sID + "'", Session["sConnSiteDB"].ToString());
                if (strResult == "1")
                {
                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else
                { return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet); }

            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
       
    }
}