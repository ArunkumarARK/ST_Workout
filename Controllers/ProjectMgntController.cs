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
    public class ProjectMgntController : Controller
    {
        private Dictionary<char, int> CharValues = null;
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: ProjectMgnt
        [SessionExpire]
        public ActionResult Index(string CustAcc, string EmpID, string SiteID) //
        {
            DataTable dtnew = new DataTable();
            //DataTable dtCloned = new DataTable();
            //appINI from xml
            //string pathVal = stINI.getINIalue("vPath", "vpWrkFold", @"C:\Applications\INI\MIS.INI");

            //select ji.JBM_AutoID, jc.CustSN, ji.JBM_ID, ji.JBM_Intrnl, ji.JBM_IM from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%BK%' 
            //For test making link
            //string strreturnURL = objDS.Encrypt("http://10.20.11.31/smarttrack/ManagerInbox.aspx", "*!%$@~&#?,:");
            //string strJBMAutoId = objDS.Encrypt("BK510", "*!%$@~&#?,:");
            //string strCustSN = objDS.Encrypt("TandF", "*!%$@~&#?,:");
            // string strEmpID = objDS.Encrypt("E000002", "*!%$@~&#?,:");
            //string strSiteID = objDS.Encrypt("L0003", "*!%$@~&#?,:");

            //SampleLink
            //http://localhost:44358/ProjectMgnt/Project?returnURL=ZRLdaunZZwRW2hkDo4H+MLt5jGpVH/IdUFvZbaFv2KTncP2W64ZW0WT1/L6u1LY1&JBMAutoId=dzVKIyBA2Ac=&CustSN=Mju5bbcjzL4=&EmpID=muhftAWj+2c===&SiteID=XFcfKokxvfo=&CustAcc=BK
            //http://localhost:44358/ProjectMgnt/Project?returnURL=ZRLdaunZZwRW2hkDo4H+MLt5jGpVH/IdUFvZbaFv2KTncP2W64ZW0WT1/L6u1LY1&JBMAutoId=kUvnTmDN2js=&CustSN=QfjUERcWfVo=&EmpID=muhftAWj+2c=&SiteID=XFcfKokxvfo=&CustAcc=BK
            if (CustAcc != null)
            {
                string strUrl = Request.Url.AbsoluteUri.ToString();

                Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";
                Session["EmpIdLogin"] = EmpID;
                Session["sCustAcc"] = CustAcc;
                Session["sSiteID"] = SiteID;
                Session["sJBMAutoID"] = "";
  
                clsCollec.getSiteDBConnection(SiteID, CustAcc);
                if (Session["sConnSiteDB"].ToString() == "")
                {
                    Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                }

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select EmpAutoId,EmpName,EmpMailId,DeptAccess,CustAccess,TeamID from JBM_EmployeeMaster Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
                    Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();
                }

                //Load Project list Items
                List<SelectListItem> lstProject = new List<SelectListItem>();

                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select  DISTINCT  jc.CustSN from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%BK%' and ji.jbm_autoid in (SELECT JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "')", Session["sConnSiteDB"].ToString());

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
                else {

                    ds = new DataSet();
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

                }
               
                ViewBag.Projectlist = lstProject;

                string strQueryFinal = "select a.JBM_AutoID,a.Short_Stage ,b.StageName from (select JBM_AutoID, Short_Stage from BK_ScheduleInfo  where JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' )a   left join   (Select distinct SD.StageName, SD.StageShortName from JBM_StageDescription SD join BK_scheduleinfo SI  on SD.StageShortName = SI.Short_Stage   where SD.IS_CustStage = 'Y' ) b on  a.Short_Stage = b.StageShortName";
                string strResult = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());

                if (strResult != "-1")
                {
                    strQueryFinal = "select a.JBM_AutoID,a.Short_Stage ,b.StageName from (select JBM_AutoID, Short_Stage from BK_ScheduleInfo  where JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and JBM_AutoID in (SELECT JBM_AutoID FROM BK_ProcessInfo Where EmpAutoId = '" + Session["EmpAutoId"].ToString() + "'))a   left join   (Select distinct SD.StageName, SD.StageShortName from JBM_StageDescription SD join BK_scheduleinfo SI  on SD.StageShortName = SI.Short_Stage   where SD.IS_CustStage = 'Y' ) b on  a.Short_Stage = b.StageShortName";
                }
                DataSet ds1 = new DataSet();
                ds1 = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                DataView view = new DataView(ds1.Tables[0]);
                DataTable distinctShort_Stage = view.ToTable(true, "StageName");
                List<SelectListItem> Useritems = new List<SelectListItem>();
                //foreach (DataRow myRow in distinctShort_Stage.Rows)
                //{
                //    Useritems.Add(new SelectListItem
                //    {
                //        Text = myRow["StageName"].ToString(),
                //        Value = myRow["StageName"].ToString()
                //    });
                //}
                Useritems.Add(new SelectListItem { Text = "Copy Editing", Value = "CE" });
                Useritems.Add(new SelectListItem { Text = "First Pages", Value = "FP" });
                Useritems.Add(new SelectListItem { Text = "Second Pages", Value = "2ndPg" });
                Useritems.Add(new SelectListItem { Text = "Final Pages", Value = "FinPag" });

                ViewBag.Userlist = Useritems;

                Session["IsLoaded"] = "Yes";
                
                try
                {
                    List<ScheduleModel> Components = new List<ScheduleModel>();
                    string strQueryFinal1 = "select CONVERT(VARCHAR(10), a.PlannedStartDate,101) as PlannedStartDate,CONVERT(VARCHAR(10), a.PlannedEndDate,101) as PlannedEndDate,CONVERT(VARCHAR(10), b.ReceivedDate,101) as ReceivedDate,CONVERT(VARCHAR(10), b.DispatchDate,101) as DispatchDate,a.JBM_AutoID,a.Short_Stage,a.CustSN,a.JBM_Intrnl,CONVERT(VARCHAR(10), a.JBM_PrinterDate,101) as JBM_PrinterDate from (select bs.PlannedStartDate, bs.PlannedEndDate, bs.JBM_AutoID, bs.Short_Stage, jc.CustSN, ji.JBM_Intrnl,ji.JBM_PrinterDate from " + Session["sCustAcc"].ToString() + "_ScheduleInfo bs  JOIN JBM_Info ji ON bs.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where bs.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ) a left join (SELECT  min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate, CI.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI JOIN BK_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID JOIN JBM_Info ji ON CI.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where  CI.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by CI.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl) b on a.Jbm_Autoid = b.JBM_AutoID and a.Short_Stage = b.Short_Stage";
                    string strResult1 = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());

                    if (strResult1 != "-1")
                    {
                        strQueryFinal1 = "select CONVERT(VARCHAR(10), a.PlannedStartDate,101) as PlannedStartDate,CONVERT(VARCHAR(10), a.PlannedEndDate,101) as PlannedEndDate,CONVERT(VARCHAR(10), b.ReceivedDate,101) as ReceivedDate,CONVERT(VARCHAR(10), b.DispatchDate,101) as DispatchDate,a.JBM_AutoID,a.Short_Stage,a.CustSN,a.JBM_Intrnl,CONVERT(VARCHAR(10), a.JBM_PrinterDate,101) as JBM_PrinterDate from (select bs.PlannedStartDate, bs.PlannedEndDate, bs.JBM_AutoID, bs.Short_Stage, jc.CustSN, ji.JBM_Intrnl,ji.JBM_PrinterDate from " + Session["sCustAcc"].ToString() + "_ScheduleInfo bs  JOIN JBM_Info ji ON bs.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where bs.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and bs.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId = '" + Session["EmpAutoId"].ToString() + "') ) a left join (SELECT  min(SI.DispatchDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate, CI.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI JOIN BK_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID JOIN JBM_Info ji ON CI.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where  CI.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and CI.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId = '" + Session["EmpAutoId"].ToString() + "') group by CI.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl) b on a.Jbm_Autoid = b.JBM_AutoID and a.Short_Stage = b.Short_Stage";
                    }
                    DataSet ds1t = new DataSet();
                    ds1t = DBProc.GetResultasDataSet(strQueryFinal1, Session["sConnSiteDB"].ToString());

                    DataView view1 = new DataView(ds1t.Tables[0]);
                    DataTable distinctJBM_AutoID = view1.ToTable(true, "JBM_AutoID", "CustSN", "JBM_Intrnl", "JBM_PrinterDate");


                    DataSet dsNew = new DataSet();

                    if (ds1t.Tables[0].Rows.Count > 0)
                    {
                        DataTable dtn = new DataTable("MyTable");
                        dtn.Columns.Add(new DataColumn("JBM_AutoID", typeof(string)));
                        dtn.Columns.Add(new DataColumn("CustSN", typeof(string)));
                        dtn.Columns.Add(new DataColumn("JBM_Intrnl", typeof(string)));
                        for (int i = 0; i < Useritems.Count; i++)
                        {
                            Boolean columnExists = false;
                            string cname = Useritems[i].Value.ToString();
                            columnExists = dtn.Columns.Contains(cname + "_PlannedStartDate");
                            if (columnExists == false)
                                dtn.Columns.Add(new DataColumn(cname + "_PlannedStartDate", typeof(string)));
                            columnExists = dtn.Columns.Contains(cname + "_PlannedEndDate");
                            if (columnExists == false)
                                dtn.Columns.Add(new DataColumn(cname + "_PlannedEndDate", typeof(string)));
                            columnExists = dtn.Columns.Contains(cname + "_ReceivedDate");
                            if (columnExists == false)
                                dtn.Columns.Add(new DataColumn(cname + "_ReceivedDate", typeof(string)));
                            columnExists = dtn.Columns.Contains(cname + "_DispatchDate");
                            if (columnExists == false)
                                dtn.Columns.Add(new DataColumn(cname + "_DispatchDate", typeof(string)));
                        }
                        dtn.Columns.Add(new DataColumn("PrinterDate", typeof(string)));
                        DataRow dr;
                        for (int intCount = 0; intCount < distinctJBM_AutoID.Rows.Count; intCount++)
                        {
                            dr = dtn.NewRow();
                            dr["JBM_AutoID"] = distinctJBM_AutoID.Rows[intCount][0].ToString();
                            dr["CustSN"] = distinctJBM_AutoID.Rows[intCount][1].ToString();
                            dr["JBM_Intrnl"] = distinctJBM_AutoID.Rows[intCount][2].ToString();
                            dr["PrinterDate"] = distinctJBM_AutoID.Rows[intCount][3].ToString();
                            dtn.Rows.Add(dr);
                        }
                        for (int j = 0; j < ds1t.Tables[0].Rows.Count; j++)
                        {
                            for (int i = 0; i < dtn.Columns.Count; i++)
                            {
                                string colname = dtn.Columns[i].ColumnName.ToString();
                                if (colname != "JBM_AutoID" && colname != "CustSN" && colname != "JBM_Intrnl" && colname != "PrinterDate")
                                {
                                    string[] colarr = colname.Split('_');
                                    string stage = colarr[0].ToString();
                                    string datename = colarr[1].ToString();
                                    if (stage == ds1t.Tables[0].Rows[j]["Short_Stage"].ToString())
                                    {
                                        for (int k = 0; k < dtn.Rows.Count; k++)
                                        {
                                            if (dtn.Rows[k]["JBM_AutoID"].ToString() == ds1t.Tables[0].Rows[j]["JBM_AutoID"].ToString())
                                            {
                                                DataRow[] rows = dtn.Select("JBM_AutoID = '" + dtn.Rows[k]["JBM_AutoID"].ToString() + "'");
                                                rows[0][colname] = ds1t.Tables[0].Rows[j][colarr[1].ToString()].ToString();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                       
                        dsNew.Tables.Add(dtn);
                    }

                   
                    //var JSONString = from a in dsNew.Tables[0].AsEnumerable()
                    //                 select new[] { CreateHREFLinkS(a[0].ToString(), "HrefLink", a[1].ToString(), a[2].ToString()), a[1].ToString(), a[6].ToString().Trim() == "[Select]" ? "" : a[6].ToString().Trim(), a[8].ToString().Trim() == "[Select]" ? "" : a[8].ToString().Trim(), a[4].ToString(), a[7].ToString(), a[5].ToString() };

                    //return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                    //return View(dsNew.Tables[0]);
                    dtnew = dsNew.Tables[0];
                    dtnew.Columns.Add("Project", typeof(HtmlString)).SetOrdinal(0);
                    for (int x = 0; x < dtnew.Rows.Count; x++)
                    {
                        HtmlString link = new HtmlString(CreateHREFLinkS(dtnew.Rows[x][1].ToString(), "HrefLink", dtnew.Rows[x][2].ToString(), dtnew.Rows[x][3].ToString()));
                        dtnew.Rows[x][0] = link;
                        DataColumnCollection columns = dtnew.Columns;
                        if (columns.Contains("CE_ReceivedDate") && columns.Contains("CE_DispatchDate"))
                        {
                            DataTable dtce = DBProc.GetResultasDataTbl("select CONVERT(VARCHAR(10), min(SI.CeDispDate),101) as ReceivedDate, CONVERT(VARCHAR(10),max(SI.CeDispDate),101) as DispatchDate  FROM  BK_ChapterInfo CI JOIN BK_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID  where  SI.RevFinStage = 'FP' and CI.JBM_AutoID = '" + dtnew.Rows[x]["JBM_AutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                            if (dtce.Rows.Count > 0)
                            {
                                dtnew.Rows[x]["CE_ReceivedDate"] = dtce.Rows[0]["ReceivedDate"].ToString();
                                dtnew.Rows[x]["CE_DispatchDate"] = dtce.Rows[0]["DispatchDate"].ToString();
                            }
                        }
                    }

                    
                }
                catch (Exception)
                {
                    return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            else {

                return View("~/Views/Shared/Error.cshtml");
            }            
            return View(dtnew);

        }
        [SessionExpire]
        public ActionResult GetProjectList(string sCustSN)
        {
            try
            {
                sCustSN = sCustSN.Replace("[Select]", "");
                string strQueryFinal = "SELECT ji.JBM_AutoID,ji.JBM_Location as Division, pm.Current_Health, pm.Current_Status, jc.CustName, jc.CustSN, CONVERT(VARCHAR(10), ji.JBM_PrinterDate,101) as JBM_PrinterDate, ji.JBM_ID, ji.Title, ji.BM_FullService, ji.KGLAccMgrName as CPM, pm.ProjectCoordInd as PM,ji.BM_DesiredPgCount as PgCount,ji.JBM_Intrnl,(ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)% 3 AS Col, (ROW_NUMBER() OVER(ORDER BY ji.JBM_AutoID) - 1)/ 3 AS Row FROM JBM_Info ji JOIN BK_ProjectManagement pm ON ji.JBM_AutoID = pm.JBM_AutoID  JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where ji.jbm_disabled = '0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
               // string strQueryFinal = "Select ji.JBM_AutoID, jc.CustSN, ji.JBM_ID, ji.JBM_Intrnl,(ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)%3 AS Col, (ROW_NUMBER() OVER (ORDER BY JBM_AutoID) -1)/3 AS Row from JBM_Info ji join JBM_CustomerMaster jc on ji.custid=jc.custid where ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                if (sCustSN != "")
                {
                    strQueryFinal  +=  " and jc.CustSN in ('" + sCustSN.Trim() + "')";
                }

                string strResult = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());

                if (strResult != "-1")
                {
                    strQueryFinal += " and ji.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "')";
                }

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());


                // Creating single column into three column view
               // DataSet dsNew = new DataSet();

                //if (ds.Tables[0].Rows.Count > 0)
                //{
                //    int ColCnt = 1;

                //    DataTable dt = new DataTable("MyTable");
                //    dt.Columns.Add(new DataColumn("Col1JBM_AIDCustSN", typeof(string)));
                //    dt.Columns.Add(new DataColumn("Col1", typeof(string)));
                //    dt.Columns.Add(new DataColumn("Col2JBM_AIDCustSN", typeof(string)));
                //    dt.Columns.Add(new DataColumn("Col2", typeof(string)));
                //    dt.Columns.Add(new DataColumn("Col3JBM_AIDCustSN", typeof(string)));
                //    dt.Columns.Add(new DataColumn("Col3", typeof(string)));

                //    DataRow dr = dt.NewRow();
                //    string strCol1Val = "";
                //    string strCol1JBM = "";
                //    string strCol2Val = "";
                //    string strCol2JBM = "";
                //    string strCol3Val = "";
                //    string strCol3JBM = "";
                //    int newRow = 0;
                //    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                //    {


                //        if (ds.Tables[0].Rows[intCount]["Col"].ToString() == "0")
                //        {
                //            strCol1Val = ds.Tables[0].Rows[intCount]["JBM_Intrnl"].ToString();
                //            strCol1JBM = ds.Tables[0].Rows[intCount]["JBM_AutoID"].ToString() + "|" + ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                //        }
                //        else if (ds.Tables[0].Rows[intCount]["Col"].ToString() == "1")
                //        {
                //            strCol2Val = ds.Tables[0].Rows[intCount]["JBM_Intrnl"].ToString();
                //            strCol2JBM = ds.Tables[0].Rows[intCount]["JBM_AutoID"].ToString() + "|" + ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                //        }
                //        else if (ds.Tables[0].Rows[intCount]["Col"].ToString() == "2")
                //        {
                //            strCol3Val = ds.Tables[0].Rows[intCount]["JBM_Intrnl"].ToString();
                //            strCol3JBM = ds.Tables[0].Rows[intCount]["JBM_AutoID"].ToString() + "|" + ds.Tables[0].Rows[intCount]["CustSN"].ToString();
                //        }


                //        if (ds.Tables[0].Rows[intCount]["Col"].ToString() == "2" && intCount != ds.Tables[0].Rows.Count)
                //        {
                //            newRow = 1;
                //        }
                //        else if (intCount == ds.Tables[0].Rows.Count - 1)
                //        {
                //            newRow = 1;
                //        }


                //        if (newRow == 1)
                //        {
                //            dr = dt.NewRow();
                //            if (strCol1JBM == "")
                //                { strCol1JBM = "|"; }
                //            if (strCol2JBM == "")
                //            { strCol2JBM = "|"; }
                //            if (strCol3JBM == "")
                //            { strCol3JBM = "|"; }
                //            dr["Col1JBM_AIDCustSN"] = strCol1JBM;
                //            dr["Col1"] = strCol1Val;
                //            dr["Col2JBM_AIDCustSN"] = strCol2JBM;
                //            dr["Col2"] = strCol2Val;
                //            dr["Col3JBM_AIDCustSN"] = strCol3JBM;
                //            dr["Col3"] = strCol3Val;

                //            dt.Rows.Add(dr);

                //            strCol1Val = "";
                //            strCol2Val = "";
                //            strCol3Val = "";
                //            strCol1JBM = "";
                //            strCol2JBM = "";
                //            strCol3JBM = "";
                //            newRow = 0;
                //        }


                //    }
                //    dsNew.Tables.Add(dt);
                //}
                
                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {CreateHREFLink(a[0].ToString(),"HrefLink", a[5].ToString(), a[13].ToString()),a[5].ToString(),a[1].ToString(),a[8].ToString(),a[10].ToString(),a[11].ToString().Trim()=="[Select]"?"":a[11].ToString().Trim(),a[12].ToString(),a[9].ToString().Trim()!=""? a[9].ToString().Trim()=="0"?"FullService":"Comp":"",a[3].ToString(),a[2].ToString().Trim()=="[Select]"?"":a[2].ToString().Trim(),a[6].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                //CreateAIBtn(a[7].ToString()
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetRemarksList()
        {
            try
            {
                string strQueryFinal = "SELECT si.AutoArtID, si.Instruction,si.InstDate,ji.JBM_ID,ji.DocketNo,pm.BusinessUnit,pm.DesignDesc,pm.ProductionLead,ji.JBM_Location,jc.CustSN,ji.JBM_Intrnl FROM BK_ProjectManagement pm JOIN BK_SplInstructions si ON pm.JBM_AutoID = si.AutoArtID JOIN JBM_Info ji ON pm.JBM_AutoID = ji.JBM_AutoID JOIN  JBM_CustomerMaster jc ON ji.CustID = jc.CustID where exists(select AutoArtID from (select AutoArtID, max(InstDate) as maxdate from BK_SplInstructions group by AutoArtID) sm  where sm.AutoArtID = si.AutoArtID and sm.maxdate = si.InstDate) and ji.jbm_disabled='0' and ji.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ";
                string strResult = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());

                if (strResult != "-1")
                {
                    strQueryFinal += " and ji.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "')";
                }
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {CreateHREFLinkR(a[0].ToString(),"HrefLink", a[9].ToString(), a[10].ToString()),a[1].ToString(),a[6].ToString().Trim()=="[Select]"?"":a[6].ToString().Trim(),a[8].ToString().Trim()=="[Select]"?"":a[8].ToString().Trim(),a[4].ToString(),a[7].ToString(),a[5].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        //public ActionResult GetScheduleList()
        //{
        //    try
        //    {
        //        List<ScheduleModel> Components = new List<ScheduleModel>();
        //        string strQueryFinal = "select CONVERT(VARCHAR(10), a.PlannedStartDate,110) as PlannedStartDate,CONVERT(VARCHAR(10), a.PlannedEndDate,110) as PlannedEndDate,CONVERT(VARCHAR(10), b.ReceivedDate,110) as ReceivedDate,CONVERT(VARCHAR(10), b.DispatchDate,110) as DispatchDate,a.JBM_AutoID,a.Short_Stage,a.CustSN,a.JBM_Intrnl from (select bs.PlannedStartDate, bs.PlannedEndDate, bs.JBM_AutoID, bs.Short_Stage, jc.CustSN, ji.JBM_Intrnl from " + Session["sCustAcc"].ToString() + "_ScheduleInfo bs  JOIN JBM_Info ji ON bs.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where bs.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' ) a left join (SELECT  min(SI.ReceivedDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate, CI.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI JOIN BK_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID JOIN JBM_Info ji ON CI.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where  CI.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%'  group by CI.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl) b on a.Jbm_Autoid = b.JBM_AutoID and a.Short_Stage = b.Short_Stage";
        //        string strResult = DBProc.GetResultasString("SELECT top 1 JBM_AUTOID FROM BK_ProcessInfo Where EmpAutoId='" + Session["EmpAutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());

        //        if (strResult != "-1")
        //        {
        //            strQueryFinal = "select CONVERT(VARCHAR(10), a.PlannedStartDate,110) as PlannedStartDate,CONVERT(VARCHAR(10), a.PlannedEndDate,110) as PlannedEndDate,CONVERT(VARCHAR(10), b.ReceivedDate,110) as ReceivedDate,CONVERT(VARCHAR(10), b.DispatchDate,110) as DispatchDate,a.JBM_AutoID,a.Short_Stage,a.CustSN,a.JBM_Intrnl from (select bs.PlannedStartDate, bs.PlannedEndDate, bs.JBM_AutoID, bs.Short_Stage, jc.CustSN, ji.JBM_Intrnl from " + Session["sCustAcc"].ToString() + "_ScheduleInfo bs  JOIN JBM_Info ji ON bs.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where bs.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and bs.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId = '" + Session["EmpAutoId"].ToString() + "') ) a left join (SELECT  min(SI.ReceivedDate) as ReceivedDate, max(SI.DispatchDate) as DispatchDate, CI.JBM_AutoID, SI.RevFinStage as Short_Stage, jc.CustSN, ji.JBM_Intrnl FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI JOIN BK_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID JOIN JBM_Info ji ON CI.JBM_AutoID = ji.JBM_AutoID JOIN JBM_CustomerMaster jc ON ji.CustID = jc.CustID where  CI.JBM_AutoID like '%" + Session["sCustAcc"].ToString() + "%' and CI.JBM_AutoID in (SELECT JBM_AutoID FROM " + Session["sCustAcc"].ToString() + "_ProcessInfo Where EmpAutoId = '" + Session["EmpAutoId"].ToString() + "') group by CI.JBM_AutoID,SI.RevFinStage,jc.CustSN,ji.JBM_Intrnl) b on a.Jbm_Autoid = b.JBM_AutoID and a.Short_Stage = b.Short_Stage";
        //        }
        //        DataSet ds = new DataSet();
        //        ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

        //        DataView view = new DataView(ds.Tables[0]);
        //        DataTable distinctJBM_AutoID = view.ToTable(true, "JBM_AutoID", "CustSN", "JBM_Intrnl");          


        //        DataSet dsNew = new DataSet();

        //        if (ds.Tables[0].Rows.Count > 0)
        //        {
        //            DataTable dtn = new DataTable("MyTable");
        //            dtn.Columns.Add(new DataColumn("JBM_AutoID", typeof(HtmlString)));
        //            dtn.Columns.Add(new DataColumn("CustSN", typeof(string)));
        //            dtn.Columns.Add(new DataColumn("JBM_Intrnl", typeof(string)));
        //            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
        //            {
        //                Boolean columnExists = false;
        //                string cname = ds.Tables[0].Rows[i][5].ToString();
        //                columnExists = dtn.Columns.Contains(cname + "_PlannedStartDate");
        //                if (columnExists == false)
        //                    dtn.Columns.Add(new DataColumn(cname + "_PlannedStartDate", typeof(string)));
        //                columnExists = dtn.Columns.Contains(cname + "_PlannedEndDate");
        //                if (columnExists == false)
        //                    dtn.Columns.Add(new DataColumn(cname + "_PlannedEndDate", typeof(string)));
        //                columnExists = dtn.Columns.Contains(cname + "_ReceivedDate");
        //                if (columnExists == false)
        //                    dtn.Columns.Add(new DataColumn(cname + "_ReceivedDate", typeof(string)));
        //                columnExists = dtn.Columns.Contains(cname + "_DispatchDate");
        //                if (columnExists == false)
        //                    dtn.Columns.Add(new DataColumn(cname + "_DispatchDate", typeof(string)));
        //            }
        //            DataRow dr;
        //            for (int intCount = 0; intCount < distinctJBM_AutoID.Rows.Count; intCount++)
        //            {
        //                dr = dtn.NewRow();
        //                dr["JBM_AutoID"] = distinctJBM_AutoID.Rows[intCount][0].ToString();
        //                dr["CustSN"] = distinctJBM_AutoID.Rows[intCount][1].ToString();
        //                dr["JBM_Intrnl"] = distinctJBM_AutoID.Rows[intCount][2].ToString();
        //                dtn.Rows.Add(dr);
        //            }
        //            for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
        //            {
        //                for (int i = 0; i < dtn.Columns.Count; i++)
        //                {
        //                    string colname = dtn.Columns[i].ColumnName.ToString();
        //                    if (colname != "JBM_AutoID" && colname != "CustSN" && colname != "JBM_Intrnl")
        //                    {
        //                        string[] colarr = colname.Split('_');
        //                        string stage = colarr[0].ToString();
        //                        string datename = colarr[1].ToString();
        //                        if (stage == ds.Tables[0].Rows[j]["Short_Stage"].ToString())
        //                        {
        //                            for (int k = 0; k < dtn.Rows.Count; k++)
        //                            {
        //                                if (dtn.Rows[k]["JBM_AutoID"].ToString() == ds.Tables[0].Rows[j]["JBM_AutoID"].ToString())
        //                                {
        //                                    DataRow[] rows = dtn.Select("JBM_AutoID = '" + dtn.Rows[k]["JBM_AutoID"].ToString() + "'");
        //                                    rows[0][colname] = ds.Tables[0].Rows[j][colarr[1].ToString()].ToString();
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            dsNew.Tables.Add(dtn);
        //        }


        //        //var JSONString = from a in dsNew.Tables[0].AsEnumerable()
        //        //                 select new[] { CreateHREFLinkS(a[0].ToString(), "HrefLink", a[1].ToString(), a[2].ToString()), a[1].ToString(), a[6].ToString().Trim() == "[Select]" ? "" : a[6].ToString().Trim(), a[8].ToString().Trim() == "[Select]" ? "" : a[8].ToString().Trim(), a[4].ToString(), a[7].ToString(), a[5].ToString() };

        //        //return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
        //        //return View(dsNew.Tables[0]);
        //        return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception)
        //    {
        //        return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
        //    }
           
        //}
        public string CreateHREFLink(string uniqueID, string strType, string strCustSN, string strJBMIntrnl)
        {
            string formControl = string.Empty;
            try
            {
                if (strCustSN != "")
                {
                    string strUrl = "";
                    strUrl = Request.Url.AbsoluteUri.ToString().Replace("GetProjectList", "Project");
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
        public string CreateHREFLinkR(string uniqueID, string strType, string strCustSN, string strJBMIntrnl)
        {
            string formControl = string.Empty;
            try  
            {
                if (strCustSN != "")
                {
                    string strUrl = "";
                    strUrl = Request.Url.AbsoluteUri.ToString().Replace("GetRemarksList", "Project");
                    GlobalVariables.strreturnURL = objDS.Encrypt(Session["returnURL"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strJBMAutoId = objDS.Encrypt(uniqueID, "*!%$@~&#?,:");
                    GlobalVariables.strCustSN = objDS.Encrypt(strCustSN, "*!%$@~&#?,:");
                    GlobalVariables.strEmpID = objDS.Encrypt(Session["EmpIdLogin"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strSiteID = objDS.Encrypt(Session["sSiteID"].ToString(), "*!%$@~&#?,:");
                    GlobalVariables.strCustAcc = Session["sCustAcc"].ToString();

                    string strNavigateURL = strUrl + "?returnURL=" + GlobalVariables.strreturnURL + "&JBMAutoId=" + GlobalVariables.strJBMAutoId + "&CustSN=" + GlobalVariables.strCustSN + "&EmpID=" + GlobalVariables.strEmpID + "&SiteID=" + GlobalVariables.strSiteID + "&CustAcc=" + GlobalVariables.strCustAcc;
                    //sample : http://localhost:44358/ProjectMgnt/Project?returnURL=BZ1AAr3qrknHGTo5xB1Ue+2kZfivafi9yUc0/rq3uYJtA1/8t+0BfFy5cKCxz9tzcJwU+NjeuUAGdCC9rRvbuTpQvw4u9J0iKv9RWoye7oWPIh4c/aqclY10cVkPrpzulTsLpQqJHC4aMZqDLSrhfQ==&JBMAutoId=i4omwZDFxHE+Z245BgWtmA==&CustSN=5M5zCkmHL9kfBOZhnRXQAQ==&EmpID=1DyHJApHKbnzBP3CVO3/rg==&SiteID=focrQ+OtCm9E7Fy5buTZIg==&CustAcc=BK
                    formControl = "<a href='" + strNavigateURL + "'>"  + strJBMIntrnl + "</a>";
                }
                


                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        public string CreateHREFLinkS(string uniqueID, string strType, string strCustSN, string strJBMIntrnl)
        {
            string formControl = string.Empty;
            try
            {
                if (strCustSN != "")
                {
                    string strUrl = "";
                    strUrl = Request.Url.AbsoluteUri.ToString().Split('?')[0].ToString();
                    strUrl = strUrl.Replace("/Index", "/Project");
                    //'strUrl = Request.Url.AbsoluteUri.ToString().Replace("Index", "Project");
                    //strUrl = Request.UrlReferrer.AbsoluteUri.ToString() + "ProjectMgnt/Project";
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
        public ActionResult Project(string returnURL, string JBMAutoId, string CustSN, string EmpID, string SiteID, string CustAcc, string AccessRights)
        {
             HtmlHelper.ClientValidationEnabled = false;
            try
            {

                //gen.WriteLog("App initialized...");
                //gen.WriteLog("ReturnURL: " + returnURL);
                //gen.WriteLog("JBM AutoID: " + JBMAutoId);
                //gen.WriteLog("Cust SN: " + CustSN);
                //gen.WriteLog("Emp ID: " + EmpID);
                //gen.WriteLog("Site ID: " + SiteID);
                //gen.WriteLog("Cust Acc: " + CustAcc);

                //To get query string details at initial loading
                if (JBMAutoId != null)
                {
                    Session["IsLoaded"] = "Yes";
                    //GlobalVariables.strreturnURL = objDS.Decrypt(returnURL, "*!%$@~&#?,:");
                    //GlobalVariables.strJBMAutoId = objDS.Decrypt(JBMAutoId, "*!%$@~&#?,:");
                    //GlobalVariables.strCustSN = objDS.Decrypt(CustSN, "*!%$@~&#?,:");
                    //GlobalVariables.strEmpID = objDS.Decrypt(EmpID, "*!%$@~&#?,:");
                    //GlobalVariables.strSiteID = objDS.Decrypt(SiteID, "*!%$@~&#?,:");
                    //GlobalVariables.strCustAcc = CustAcc;

                    Session["sJBMAutoID"] = objDS.Decrypt(JBMAutoId, "*!%$@~&#?,:");
                    Session["sCustSN"] = objDS.Decrypt(CustSN, "*!%$@~&#?,:");
                    Session["EmpIdLogin"] = objDS.Decrypt(EmpID, "*!%$@~&#?,:");
                    Session["sSiteID"] = objDS.Decrypt(SiteID, "*!%$@~&#?,:");
                    Session["sCustAcc"] = CustAcc;

                    // GlobalVariables.strJBMAutoId = "BK050"; // for test
                    //string strUrl = objDS.Decrypt(returnURL, "*!%$@~&#?,:");   //ViewBag.strReturnURL
                    Session["strReturnURL"] = objDS.Decrypt(returnURL, "*!%$@~&#?,:");  //strUrl.Replace("ProjectTrack/Index", "ProjectTrack/ProjectTracking").Split('?')[0].ToString();

                    if (Session["sSiteID"].ToString() == "L0002")
                    {
                        return View("~/Views/Shared/Error.cshtml");
                    }
                    else if (Session["sSiteID"].ToString() == "L0003")
                    {
                        //Session["sConnSiteDB"].ToString() = "dbConnSmartND"; //"dbConnSmartNDTest" //BK013
                        clsCollec.getSiteDBConnection(Session["sSiteID"].ToString(), CustAcc);
                        if (Session["sConnSiteDB"].ToString() == "")
                        {
                            Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                        }
                    }


                        if (Session["sCustSN"].ToString() == "" && Session["sJBMAutoID"].ToString() == "")
                    {
                        return Redirect(Session["strReturnURL"].ToString());
                    }
                }

                

                DataSet ds1 = new DataSet();
                ds1 = DBProc.GetResultasDataSet("Select EmpAutoId, DeptCode,EmpName from JBM_EmployeeMaster WHERE EmpLogin = '" + Session["EmpIdLogin"].ToString() + "';", Session["sConnSiteDB"].ToString());
                if (ds1.Tables[0].Rows.Count > 0)
                {
                    //GlobalVariables.strDeptCode = ds1.Tables[0].Rows[0]["DeptCode"].ToString().Trim();
                    //GlobalVariables.strEmpAutoID =  ds1.Tables[0].Rows[0]["EmpAutoId"].ToString().Trim();
                    //GlobalVariables.strEmpName = ds1.Tables[0].Rows[0]["EmpName"].ToString().Trim(); ;
                    Session["DeptCode"] = ds1.Tables[0].Rows[0]["DeptCode"].ToString().Trim();
                    Session["EmpAutoId"] = ds1.Tables[0].Rows[0]["EmpAutoId"].ToString().Trim();
                    Session["EmpName"] = ds1.Tables[0].Rows[0]["EmpName"].ToString().Trim(); ;
                }

                if (Session["RoleID"].ToString()== "102")
                {
                    GlobalVariables.strAccessRights = "Customer";
                    ViewBag.AccessRight = "Customer";
                    Session["AccessRights"] = "Customer";
                }
                else if (Session["RoleID"].ToString() == "103")
                {
                    GlobalVariables.strAccessRights = "LPM";
                    Session["AccessRights"] = "LPM";
                }
                else if (Session["RoleID"].ToString() == "104")
                {
                    GlobalVariables.strAccessRights = "PM";
                    Session["AccessRights"] = "PM";
                }

                if (Session["IsLoaded"].ToString() == null)
                {
                    return Redirect("https://smarttrack.kwglobal.com/SmartTrack-ND");
                }

                /// For temp update for existing data dump

                DataSet dsTemp = new DataSet();
                dsTemp = DBProc.GetResultasDataSet("Select Jbm_Autoid,IntrnlID,ChapterId,NumofMSP,Castoff,StartPage,EndPage,ActualPages,'' as [Deviation],AutoArtID,Active from " + Session["sCustAcc"].ToString() + "_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString()); //BK022

                for (int intCount = 0; intCount < dsTemp.Tables[0].Rows.Count; intCount++)
                {
                    string sAutoArtID = dsTemp.Tables[0].Rows[intCount]["AutoArtID"].ToString();
                    string sStartPage = dsTemp.Tables[0].Rows[intCount]["StartPage"].ToString();
                    string sEndPage = dsTemp.Tables[0].Rows[intCount]["EndPage"].ToString();
                    string sActualPage = "";
                    sActualPage = CalcActualDeviation("Actual","", sStartPage, sEndPage, "No");

                    //To update 
                    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET ActualPages='" + sActualPage + "' WHERE AutoArtID='" + sAutoArtID + "'", Session["sConnSiteDB"].ToString());

                }

                ////End







                ProjectMgntModels model = new ProjectMgntModels();

                string strQuery = "Select JI.JBM_ID,JI.NoofChapters,JI.JBM_Location,JI.Title,JI.BM_Author,JI.BM_ISBNnumber13,JI.BM_ISBN10number,PM.Ebook_ISBN as [eISBN],JI.JBM_Platform,JI.JBM_PrinterDate,JI.CopyrightOwner,JI.JBM_Trimsize,(Select Sum (case when NumofMSP is null then 0 else NumofMSP end) as [MssPages] From  BK_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "')  as [MssPages],(Select Sum (case when Castoff is null then 0 else Castoff end) as [Castoff] From  BK_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "') as [BM_CastOffPgCount],JI.Docketno,JI.KGLAccMgrName,JI.IndexerName,JI.BM_FullService,PM.ProjectCoordInd,PM.Edition,PM.PONumber,PM.BusinessUnit,PM.Cenveo_Facility,PM.ProjectManagerUS, PM.DesignLead, PM.Copyeditor, PM.ProductionLead, PM.Proofreader,PM.Current_Status,PM.Current_Health,PM.Instock,PM.SignatureSize,PM.Design,PM.DesignDesc,PM.Color,PM.FilesPath,(Select top 1 spl.Instruction from " + Session["sCustAcc"].ToString() + "_SplInstructions spl where spl.AutoArtID='" + Session["sJBMAutoID"].ToString() + "' ORDER BY CONVERT(DateTime, spl.InstDate,101)  DESC) as [Remarks],PM.Components_Awaited,PM.Last_Review,(select sum (case when [ActualPages] is null then 0 else [ActualPages] end) as [ActualPages] From  BK_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "') as [ActualPages],JI.BM_DesiredPgCount as TargetCount,PM.Launch,PM.misc1,PM.misc2,PM.misc3,PM.misc4,PM.HardBack_ISBN,PM.OverallCost,PM.AdditionalCost,PM.TeamLead,PM.StageID from JBM_Info JI JOIN " + Session["sCustAcc"].ToString() +  "_ProjectManagement PM ON JI.JBM_AutoID = PM.JBM_AutoID where JI.JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'";
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    model.Project = Convert.ToString(ds.Tables[0].Rows[0]["JBM_ID"].ToString().Trim());
                    model.Title = Convert.ToString(ds.Tables[0].Rows[0]["Title"].ToString().Trim());
                    model.Author = Convert.ToString(ds.Tables[0].Rows[0]["BM_Author"].ToString().Trim());
                    model.ISBN13 = Convert.ToString(ds.Tables[0].Rows[0]["BM_ISBNnumber13"].ToString().Trim());
                    model.ISBN10 = Convert.ToString(ds.Tables[0].Rows[0]["BM_ISBN10number"].ToString().Trim());
                    model.eISBN = Convert.ToString(ds.Tables[0].Rows[0]["eISBN"].ToString().Trim());
                    model.HardBack_ISBN = Convert.ToString(ds.Tables[0].Rows[0]["HardBack_ISBN"].ToString().Trim());
                    model.PagingApp = Convert.ToString(ds.Tables[0].Rows[0]["JBM_Platform"].ToString().Trim());
                    model.PrinterDate = ds.Tables[0].Rows[0]["JBM_PrinterDate"].ToString().Trim() != "" ? Convert.ToDateTime(ds.Tables[0].Rows[0]["JBM_PrinterDate"].ToString().Trim()) : (DateTime?)null;
                    int flag = 0;
                    DataSet dsOther = new DataSet();
                    dsOther = DBProc.GetResultasDataSet("Select TrimSize from [dbo].[JBM_TrimSize] WHERE Status=1", Session["sConnSiteDB"].ToString());
                    for(int h=0;h<dsOther.Tables[0].Rows.Count;h++)
                    {
                        string trim = ds.Tables[0].Rows[0]["JBM_Trimsize"].ToString().Trim();
                        string othertrim= dsOther.Tables[0].Rows[h]["Trimsize"].ToString().Trim();
                        if(trim == othertrim)
                        {
                            flag = 1;
                        }
                    }
                    if (flag == 1)
                    {
                        model.TrimSize = Convert.ToString(ds.Tables[0].Rows[0]["JBM_Trimsize"].ToString().Trim());
                    }
                    else
                    {
                        model.TrimSize = "Other";
                        model.OtherTrimSize = Convert.ToString(ds.Tables[0].Rows[0]["JBM_Trimsize"].ToString().Trim());
                    }

                    model.CopyRight = Convert.ToString(ds.Tables[0].Rows[0]["CopyrightOwner"].ToString().Trim());
                    model.MsCount = ds.Tables[0].Rows[0]["MssPages"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["MssPages"].ToString().Trim()) : 0;
                    model.CastOffCount = ds.Tables[0].Rows[0]["BM_CastOffPgCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["BM_CastOffPgCount"].ToString().Trim()) : 0;
                    model.Docket = Convert.ToString(ds.Tables[0].Rows[0]["Docketno"].ToString().Trim());

                    model.Service = Convert.ToString(ds.Tables[0].Rows[0]["BM_FullService"].ToString().Trim());

                    model.ProjectMangerUS = Convert.ToString(ds.Tables[0].Rows[0]["ProjectCoordInd"].ToString().Trim());
                    model.Edition = Convert.ToString(ds.Tables[0].Rows[0]["Edition"].ToString().Trim());
                    model.PO = Convert.ToString(ds.Tables[0].Rows[0]["PONumber"].ToString().Trim());
                    model.LeadProjectManager = Convert.ToString(ds.Tables[0].Rows[0]["KGLAccMgrName"].ToString().Trim());
                    model.CustomerPM = Convert.ToString(ds.Tables[0].Rows[0]["ProjectManagerUS"].ToString().Trim());
                    model.DesignLead = Convert.ToString(ds.Tables[0].Rows[0]["DesignLead"].ToString().Trim());
                    model.CopyEditor = Convert.ToString(ds.Tables[0].Rows[0]["Copyeditor"].ToString().Trim());
                    model.ProductionLead = Convert.ToString(ds.Tables[0].Rows[0]["ProductionLead"].ToString().Trim());
                    model.ProofReader = Convert.ToString(ds.Tables[0].Rows[0]["Proofreader"].ToString().Trim());
                    model.Indexer = Convert.ToString(ds.Tables[0].Rows[0]["IndexerName"].ToString().Trim());
                    model.SubDivision = Convert.ToString(ds.Tables[0].Rows[0]["BusinessUnit"].ToString().Trim());
                    model.KGLFacility = Convert.ToString(ds.Tables[0].Rows[0]["Cenveo_Facility"].ToString().Trim());
                    model.CurrentStatus = Convert.ToString(ds.Tables[0].Rows[0]["Current_Status"].ToString().Trim());
                    model.CurrentHealth = Convert.ToString(ds.Tables[0].Rows[0]["Current_Health"].ToString().Trim());
                    model.Instock = ds.Tables[0].Rows[0]["Instock"].ToString().Trim() != "" ? Convert.ToDateTime(ds.Tables[0].Rows[0]["Instock"].ToString().Trim()) : (DateTime?)null;
                    model.Design = Convert.ToString(ds.Tables[0].Rows[0]["Design"].ToString().Trim());
                    model.DesignName = Convert.ToString(ds.Tables[0].Rows[0]["DesignDesc"].ToString().Trim());
                    model.Color = Convert.ToString(ds.Tables[0].Rows[0]["Color"].ToString().Trim());
                    model.Files = Convert.ToString(ds.Tables[0].Rows[0]["FilesPath"].ToString().Trim());
                    model.Remarks = Convert.ToString(ds.Tables[0].Rows[0]["Remarks"].ToString().Trim());
                    model.ComponentsAwaited = Convert.ToString(ds.Tables[0].Rows[0]["Components_Awaited"].ToString().Trim());
                    model.LastReview = ds.Tables[0].Rows[0]["Last_Review"].ToString().Trim() != "" ? Convert.ToDateTime(ds.Tables[0].Rows[0]["Last_Review"].ToString().Trim()) : (DateTime?)null;

                    model.SignatureSize = ds.Tables[0].Rows[0]["SignatureSize"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["SignatureSize"].ToString().Trim()) : 0;
                    model.TargetCount = ds.Tables[0].Rows[0]["TargetCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["TargetCount"].ToString().Trim()) : 0;
                    model.Currentcount = ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim()) : 0;
                    
                    // Stopped this logic after discussed with Sujeet on 08th Oct 21
                    //if (ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim() == "0" || ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim() == "")
                    //{
                    //    model.Currentcount = ds.Tables[0].Rows[0]["TargetCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["TargetCount"].ToString().Trim()) : 0;
                    //}
                    //else {
                    //    model.Currentcount = ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim()) : 0;
                    //}
                    
                    model.ChapterCount = ds.Tables[0].Rows[0]["NoofChapters"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["NoofChapters"].ToString().Trim()) : 0;
                    model.CustomerDivision = Convert.ToString(ds.Tables[0].Rows[0]["JBM_Location"].ToString().Trim());
                    model.Completion = ds.Tables[0].Rows[0]["JBM_PrinterDate"].ToString().Trim() != "" ? Convert.ToDateTime(ds.Tables[0].Rows[0]["JBM_PrinterDate"].ToString().Trim()) : (DateTime?)null;


                    model.Launch = ds.Tables[0].Rows[0]["Launch"].ToString().Trim() != "" ? Convert.ToDateTime(ds.Tables[0].Rows[0]["Launch"].ToString().Trim()) : (DateTime?)null;
                    model.Deviation = (model.Currentcount) - (model.CastOffCount);
                    if(model.Currentcount>0)
                    model.DeviationPercent = ((Convert.ToDouble(model.Deviation) / Convert.ToDouble(model.CastOffCount)) * 100).ToString("#.00");

                    model.Misc1 = Convert.ToString(ds.Tables[0].Rows[0]["Misc1"].ToString().Trim());
                    model.Misc2 = Convert.ToString(ds.Tables[0].Rows[0]["Misc2"].ToString().Trim());
                    model.Misc3 = Convert.ToString(ds.Tables[0].Rows[0]["Misc3"].ToString().Trim());
                    model.Misc4 = Convert.ToString(ds.Tables[0].Rows[0]["Misc4"].ToString().Trim());
                    model.OverallCost = Convert.ToString(ds.Tables[0].Rows[0]["OverallCost"].ToString().Trim());
                    model.AdditionalCost = Convert.ToString(ds.Tables[0].Rows[0]["AdditionalCost"].ToString().Trim());
                    model.TeamLead = Convert.ToString(ds.Tables[0].Rows[0]["TeamLead"].ToString().Trim());

                    Session["SignatureSize"] = model.SignatureSize;

                    Session["ProjectID"] = model.Project;
                    Session["CustomerName"] = Session["sCustSN"].ToString();
                    Session["DocketNo"] = model.Docket;
                    string StageID= ds.Tables[0].Rows[0]["StageID"].ToString().Trim();
                    if (StageID != "")
                        Session["StageID"] = "%" + ds.Tables[0].Rows[0]["StageID"].ToString().Trim() + "%";
                    else
                        Session["StageID"] = "";
                    //Load PagingApp list Items
                    List<SelectListItem> items = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select PlatformID,PlatformDesc from JBM_Platform", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            items.Add(new SelectListItem
                            {
                                Text = myRow["PlatformDesc"].ToString(),
                                Value = myRow["PlatformID"].ToString()
                            });
                        }
                    }
                    ViewBag.PagAppList = items;

                    //Load TrimSize list Items
                    List<SelectListItem> itemsTrim = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select TrimSize from [dbo].[JBM_TrimSize] WHERE Status=1", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            itemsTrim.Add(new SelectListItem
                            {
                                Text = myRow["TrimSize"].ToString(),
                                Value = myRow["TrimSize"].ToString()
                            });
                        }
                    }
                    ViewBag.TrimSizeList = itemsTrim;

                    //Load Lead Project Manager list Items
                    List<SelectListItem> itemsLeadProjectManager = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("select EmpAutoID,EmpLogin,EmpName from [dbo].[JBM_EmployeeMaster] where roleid='103' order by empname asc", Session["sConnSiteDB"].ToString());
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
                        itemsKGLFacility.Add(new SelectListItem{ Text = "Richmond, VA", Value = "Richmond, VA" });
                        //itemsKGLFacility.Add(new SelectListItem { Text = "Columbia, MD", Value = "Columbia, MD" }); // Ben requested
                        //itemsKGLFacility.Add(new SelectListItem { Text = "Fort Washington, PA", Value = "Fort Washington, PA" });
                        itemsKGLFacility.Add(new SelectListItem { Text = "London, UK", Value = "London, UK" });
                    }
                    ViewBag.KGLFacilityList = itemsKGLFacility;

                    //Production Lead
                    List<SelectListItem> lstProductionLead = new List<SelectListItem>();
                    ds = new DataSet();
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


                    return View(model);
                }
                else
                {
                    return Redirect(GlobalVariables.strreturnURL);
                }
                
            }
            catch(Exception ex)
            {
                return View();
            }
        }

        [HttpPost]		  
								   
        public JsonResult UpdateOthers(ProjectMgntModels obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            //if (obj.CustomerDivision.Trim() == "HSS")
            //    obj.CustomerDivision = "UK";
            //else
            //    obj.CustomerDivision = "US";
            obj.Remarks = string.IsNullOrEmpty(obj.Remarks) ? "" : obj.Remarks.ToString().Replace("'", "''").ToString();
            obj.ComponentsAwaited = string.IsNullOrEmpty(obj.ComponentsAwaited) ? "" : obj.ComponentsAwaited.ToString().Replace("'", "''").ToString();
            obj.Misc1 = string.IsNullOrEmpty(obj.Misc1) ? "" : obj.Misc1.ToString().Replace("'", "''").ToString();
            obj.Misc2 = string.IsNullOrEmpty(obj.Misc2) ? "" : obj.Misc2.ToString().Replace("'", "''").ToString();
            obj.Misc3 = string.IsNullOrEmpty(obj.Misc3) ? "" : obj.Misc3.ToString().Replace("'", "''").ToString();
            obj.Misc4 = string.IsNullOrEmpty(obj.Misc4) ? "" : obj.Misc4.ToString().Replace("'", "''").ToString();

            con.Open();
            SqlCommand cmdJI = new SqlCommand("update JBM_Info set JBM_Location='" + obj.CustomerDivision + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            SqlCommand cmdPM = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_ProjectManagement set Cenveo_Facility='" + obj.KGLFacility + "',FilesPath='" + obj.Files + "',Remarks='" + obj.Remarks + "',Components_Awaited='" + obj.ComponentsAwaited + "',misc1='" + obj.Misc1 + "',misc2='" + obj.Misc2 + "',misc3='" + obj.Misc3 + "',misc4='" + obj.Misc4 + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            //SqlCommand cmd = new SqlCommand("update JBM_Info set JBM_Location='"+ obj.CustomerDivision+ "' where JBM_AutoID='BK510'", con);
            cmdJI.ExecuteNonQuery();
            cmdPM.ExecuteNonQuery();
            //cmd.ExecuteNonQuery();
            con.Close();
            return Json(JsonRequestBehavior.AllowGet);
								  
												 
        }
        [HttpPost]
									
        public JsonResult UpdateProject(ProjectMgntModels obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            obj.CurrentStatus = obj.CurrentStatus.ToString().Replace("[Select]", "").Replace(" ", " ");
            obj.CurrentHealth = obj.CurrentHealth.ToString().Replace("[Select]", "").Replace(" ", " ");
            obj.PagingApp = obj.PagingApp.ToString().Replace("[Select]", "");

            obj.Project = string.IsNullOrEmpty(obj.Project) ? "" : obj.Project.ToString().Replace("'", "''").ToString();
            obj.Title = string.IsNullOrEmpty(obj.Title) ? "" : obj.Title.ToString().Replace("'", "''").ToString();
            obj.Author = string.IsNullOrEmpty(obj.Author) ? "" : obj.Author.ToString().Replace("'", "''").ToString();

            con.Open();
            SqlCommand cmdJI = new SqlCommand("update JBM_Info set Docketno='" + obj.Docket + "',JBM_ID='" + obj.Project + "',JBM_Intrnl='" + obj.Project + "',Title='" + obj.Title + "',BM_Author='" + obj.Author + "',BM_ISBNnumber13='" + obj.ISBN13 + "',BM_ISBN10number='" + obj.ISBN10 + "',JBM_Platform='" + obj.PagingApp + "',BM_FullService='" + obj.Service + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            SqlCommand cmdPM = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_ProjectManagement set Edition='" + obj.Edition + "',Current_Status='" + obj.CurrentStatus + "',Current_Health='" + obj.CurrentHealth + "',Ebook_ISBN='" + obj.eISBN + "',PONumber='" + obj.PO + "',HardBack_ISBN='"+obj.HardBack_ISBN+ "',OverallCost='" + obj.OverallCost + "',AdditionalCost='" + obj.AdditionalCost + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            // SqlCommand cmd = new SqlCommand("update JBM_Info set JBM_Location='"+ obj.CustomerDivision+ "' where JBM_AutoID='BK510'", con);
            cmdJI.ExecuteNonQuery();
            cmdPM.ExecuteNonQuery();
            con.Close();
            Session["ProjectID"] = obj.Project;
            return Json(JsonRequestBehavior.AllowGet);
								  
									  
			 
			 
												  
        }

        
        [HttpPost]
        [SessionExpire]
        public ActionResult UpdateSize(ProjectMgntModels obj)
        {
            Session["SignatureSize"] = obj.SignatureSize;
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            obj.Design = string.IsNullOrEmpty(obj.Design) ? "" : obj.Design.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.DesignName = string.IsNullOrEmpty(obj.DesignName) ? "" : obj.DesignName.ToString().Replace("'", "''").ToString();

            if (obj.TrimSize.Trim() == "Other")
            {
                obj.TrimSize = string.IsNullOrEmpty(obj.OtherTrimSize) ? "" : obj.OtherTrimSize.ToString();
            }

            con.Open();
            SqlCommand cmdJI = new SqlCommand("update JBM_Info set JBM_Trimsize='" + obj.TrimSize + "',BM_DesiredPgCount='" + obj.TargetCount + "',MssPages='" + obj.MsCount + "',NoofChapters='" + obj.ChapterCount + "',BM_CastOffPgCount='" + obj.CastOffCount + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            SqlCommand cmdPM = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_ProjectManagement set SignatureSize='" + obj.SignatureSize + "',Design='" + obj.Design + "',DesignDesc='" + obj.DesignName + "',Color='" + obj.Color + "',ActualPages='" + obj.Currentcount + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            // SqlCommand cmd = new SqlCommand("update JBM_Info set JBM_Location='"+ obj.CustomerDivision+ "' where JBM_AutoID='BK510'", con);
            cmdJI.ExecuteNonQuery();
            cmdPM.ExecuteNonQuery();
            con.Close();
            return Json(JsonRequestBehavior.AllowGet);
								  

											   
        }
        [HttpPost]
        [SessionExpire]
        //[AllowAnonymous]
        //[ValidateAntiForgeryToken]
        public ActionResult ScheuleUp(ProjectMgntModels obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            con.Close();
            return Json(JsonRequestBehavior.AllowGet);
        }

        [HttpPost]	
        [SessionExpire]
        public JsonResult UpdateSchedule(ProjectMgntModels obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());

            con.Open();

            string dtPrint, dtLaunch, dtInstock; //, dtCompletion;
            if (string.IsNullOrEmpty(obj.PrinterDate.ToString())){dtPrint = "null";}else { dtPrint = "'" + obj.PrinterDate + "'"; }
            if (string.IsNullOrEmpty(obj.Launch.ToString())){dtLaunch = "null";}else { dtLaunch = "'" + obj.Launch + "'"; }
            if (string.IsNullOrEmpty(obj.Instock.ToString())) { dtInstock = "null"; } else { dtInstock = "'" + obj.Instock + "'"; }
            //if (obj.Completion == null) { dtCompletion = "null"; } else { dtCompletion = "'" + obj.Completion + "'"; }

            SqlCommand cmdJI = new SqlCommand("update JBM_Info set JBM_PrinterDate=" + dtPrint + ",CopyrightOwner='" + obj.CopyRight + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            SqlCommand cmdPM = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_ProjectManagement set Instock=" + dtInstock + ",Launch=" + dtLaunch + " where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            // SqlCommand cmd = new SqlCommand("update JBM_Info set JBM_Location='"+ obj.CustomerDivision+ "' where JBM_AutoID='BK510'", con);
            cmdJI.ExecuteNonQuery();
            cmdPM.ExecuteNonQuery();
            con.Close();
            return Json(JsonRequestBehavior.AllowGet);
								  
												   
        }
        [HttpPost]				  
									
        public JsonResult UpdateTeam(ProjectMgntModels obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            obj.LeadProjectManager = obj.LeadProjectManager.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.ProjectMangerUS = obj.ProjectMangerUS.ToString().Replace("[Select]", "").Replace("'", "''").ToString();

            obj.ProductionLead = string.IsNullOrEmpty(obj.ProductionLead) ? "" : obj.ProductionLead.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.CustomerPM = string.IsNullOrEmpty(obj.CustomerPM) ? "" : obj.CustomerPM.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.Indexer = string.IsNullOrEmpty(obj.Indexer) ? "" : obj.Indexer.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.DesignLead = string.IsNullOrEmpty(obj.DesignLead) ? "" : obj.DesignLead.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.CopyEditor = string.IsNullOrEmpty(obj.CopyEditor) ? "" : obj.CopyEditor.ToString().Replace("[Select]", "").Replace("'", "''").ToString();
            obj.ProofReader = string.IsNullOrEmpty(obj.ProofReader) ? "" : obj.ProofReader.ToString().Replace("[Select]", "").Replace("'", "''").ToString();

            con.Open();
            string dtLastReview;
            if (string.IsNullOrEmpty(obj.LastReview.ToString())) { dtLastReview = "null"; } else { dtLastReview = "'" + obj.LastReview + "'"; }

            SqlCommand cmdJI = new SqlCommand("update JBM_Info set KGLAccMgrName='" + obj.LeadProjectManager + "',IndexerName='" + obj.Indexer + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            SqlCommand cmdPM = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_ProjectManagement set BusinessUnit='" + obj.SubDivision + "',ProjectCoordInd='" + obj.ProjectMangerUS + "',Last_Review=" + dtLastReview + ",ProjectManagerUS='" + obj.CustomerPM + "',DesignLead='" + obj.DesignLead + "',Copyeditor='" + obj.CopyEditor + "',ProductionLead='" + obj.ProductionLead + "',Proofreader='" + obj.ProofReader + "',TeamLead='" + obj.TeamLead + "' where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            //SqlCommand cmd = new SqlCommand("update JBM_Info set IndexerName='" + obj.Indexer+ "' where JBM_AutoID='BK510'", con);
            cmdJI.ExecuteNonQuery();
            cmdPM.ExecuteNonQuery();
            //cmd.ExecuteNonQuery();
            con.Close();
            return Json(JsonRequestBehavior.AllowGet);
								   
        }


        /// <summary>
        /// Stages Page
        /// </summary>
        /// <returns></returns>
        [SessionExpire]
        public ActionResult Stages()
        {

            DataSet ds = new DataSet();

            List<SelectListItem> stgPrimary = new List<SelectListItem>();

            string strQuery = string.Empty;
            string strCustStageGrp = string.Empty;

            if (Session["CustomerName"].ToString() != "TandF")
            {
                strCustStageGrp = " and CustStageGroup is NULL";
            }
            else
            {
                if (Regex.IsMatch(Session["StageID"].ToString(), "(PM1|PM2|PM3)", RegexOptions.IgnoreCase))
                    strCustStageGrp = " and CustStageGroup='TF'";
                else if (Regex.IsMatch(Session["StageID"].ToString(), "(CR)", RegexOptions.IgnoreCase))
                    strCustStageGrp = " and CustStageGroup='TFCR'";
            }

            strQuery = "Select StageName,StageShortName,SG.StageSeqID,SG.IsEditable from JBM_StageDescription sd join JBM_StageGroupTAT SG on sd.StageShortName=sg.StageShortID  where IS_CustStage = 'Y' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + "";

            ds = DBProc.GetResultasDataSet(strQuery + " order by sg.StageSeqID asc", Session["sConnSiteDB"].ToString());
            for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
            {
                string strStageName = ds.Tables[0].Rows[intCount]["StageName"].ToString();
                string strShortStg = ds.Tables[0].Rows[intCount]["StageShortName"].ToString();
                string strStgSeq = ds.Tables[0].Rows[intCount]["StageSeqID"].ToString();
                string IsEditable = ds.Tables[0].Rows[intCount]["IsEditable"].ToString();
                stgPrimary.Add(new SelectListItem
                {
                    Text = strStageName,
                    Value = strStgSeq + "|" + strShortStg + "|" + IsEditable
                });
            }

            //For footer
            ViewBag.PrimaryStgList = stgPrimary;

            //Static primary stages 
            DataSet ds1 = new DataSet();

            ds1 = DBProc.GetResultasDataSet("Select StageShortName,SG.StageSeqID,DefaultTAT from JBM_StageDescription sd join JBM_StageGroupTAT SG  on sd.StageShortName=sg.StageShortID WHERE StageShortName not in (Select distinct Short_Stage from " + Session["sCustAcc"].ToString() + "_ScheduleInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "') and  Is_CustStage='Y' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + "", Session["sConnSiteDB"].ToString());

            for (int intCount = 0; intCount < ds1.Tables[0].Rows.Count; intCount++)
            {
                string strStageSeqId = ds1.Tables[0].Rows[intCount]["StageShortName"].ToString().Trim();
                string strStageSeq = ds1.Tables[0].Rows[intCount]["StageSeqId"].ToString().Trim();
                string strTAT = ds1.Tables[0].Rows[intCount]["DefaultTAT"].ToString().Trim();

                string dtSchStart = "null", dtSchEnd = "null", dtRevStart = "null", dtRevEnd = "null", dtActStart = "null", dtActEnd = "null";
                string strDuplicate = "Select SeqID from " + Session["sCustAcc"].ToString() + "_ScheduleInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and SeqID='" + strStageSeq + "'";
                string strResult = DBProc.GetResultasString("If not exists (" + strDuplicate + ") INSERT INTO " + Session["sCustAcc"].ToString() + "_ScheduleInfo (JBM_AutoID,CompletionDays,PlannedStartDate, PlannedEndDate, RevisedPlanStartDate, RevisedPlanEndDate, ActualStartDate, ActualEndDate, Short_Stage, SeqID) VALUES ('" + Session["sJBMAutoID"].ToString() + "','" + strTAT + "'," + dtSchStart + "," + dtSchEnd + "," + dtRevStart + "," + dtRevEnd + "," + dtActStart + "," + dtActEnd + ",'" + strStageSeqId + "','" + strStageSeq + "')", Session["sConnSiteDB"].ToString());

            }


            ds = new DataSet();
            //For 1pp correction update min and max date taken from Pub, PR, AU and Schedule date mapping in Stage Info table
            string strQueryPrimarWithOther = "";

            strQueryPrimarWithOther = "Select SI.Jbm_Autoid,SI.ProcessID,SI.CompletionDays,SD.StageName as [PrimaryStages], CASE WHEN SI.Short_Stage = 'CE' THEN (select  min(CEDueDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') ELSE (select  min(DueDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [PlannedStartDate], CASE WHEN SI.Short_Stage = 'CE' THEN (select  max(CEDueDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') ELSE (select  max(DueDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [PlannedEndDate], CASE WHEN SI.Short_Stage = 'CE' THEN (select  min(CERevisedDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') ELSE (select  min(RevisedDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [RevisedPlanStartDate], CASE WHEN SI.Short_Stage = 'CE' THEN (select  max(CERevisedDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') ELSE (select  max(RevisedDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [RevisedPlanEndDate], CASE WHEN SI.Short_Stage = 'CE' THEN (select  min(CeDispDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') WHEN SI.Short_Stage = '1pco' THEN ((SELECT MIN(CorrMinDate) AS LastUpdateDate FROM (select PrintFinalDue  AS CorrMinDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and PrintFinalDue is not null union all select  PR_corr_Appr  AS CorrMinDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and PR_corr_Appr is not null union all select  Aut_Corr_Appr  AS CorrMinDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and Aut_Corr_Appr is not null) CorrectionDate)) ELSE (select  min(DispatchDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [ActualStartDate], CASE WHEN SI.Short_Stage = 'CE' THEN (select  max(CeDispDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage='FP') WHEN SI.Short_Stage = '1pco' THEN ((SELECT MAX(CorrMaxDate) AS LastUpdateDate FROM (select PrintFinalDue  AS CorrMaxDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and PrintFinalDue is not null union all select  PR_corr_Appr  AS CorrMaxDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and PR_corr_Appr is not null union all select  Aut_Corr_Appr  AS CorrMaxDate from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and Aut_Corr_Appr is not null) CorrectionDate)) ELSE (select  max(DispatchDate) from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active=1)) and RevFinStage=SI.Short_Stage) END AS [ActualEndDate], SI.Short_Stage,SI.SeqID,SG.IsEditable from " + Session["sCustAcc"].ToString() + "_ScheduleInfo SI JOIN JBM_StageDescription SD join JBM_StageGroupTAT SG  on sd.StageShortName=sg.StageShortID ON SI.Short_Stage = SD.StageShortName  WHERE SI.JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and SD.Is_CustStage = 'Y' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + " and SG.IsEditable is null  and SI.DeleteYN ='N' or SI.DeleteYN is null and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + " and SG.IsEditable is null ";

            // To add other stages information from ScheduleInfo tabel
            strQueryPrimarWithOther += " UNION Select SI.Jbm_Autoid,SI.ProcessID,SI.CompletionDays,SD.StageName as [PrimaryStages],PlannedStartDate,PlannedEndDate,RevisedPlanStartDate,RevisedPlanEndDate,ActualStartDate,ActualEndDate,Short_Stage,SeqID,SG.IsEditable from " + Session["sCustAcc"].ToString() + "_ScheduleInfo SI JOIN JBM_StageDescription SD join JBM_StageGroupTAT SG  on sd.StageShortName=sg.StageShortID ON SI.Short_Stage = SD.StageShortName  where  SG.IsEditable is not null and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + " order by SeqID;"; //SI.DeleteYN ='N' or SI.DeleteYN is null and
            
            ds = DBProc.GetResultasDataSet(strQueryPrimarWithOther, Session["sConnSiteDB"].ToString());

            List<StagesModel> Components = new List<StagesModel>();
            DateTime dateEmpty = new DateTime(1900, 1, 1);
            for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
            {

                Dictionary<string, DateTime> StartEnd = new Dictionary<string, DateTime>();

                if (ds.Tables[0].Rows[intCount]["PlannedStartDate"].ToString() != "")
                {
                    StartEnd.Add("PlannedStartDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["PlannedStartDate"])); //
                }
                else
                {
                    StartEnd.Add("PlannedStartDate", dateEmpty);
                }
                if (ds.Tables[0].Rows[intCount]["PlannedEndDate"].ToString() != "")
                {
                    StartEnd.Add("PlannedEndDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["PlannedEndDate"]));
                }
                else
                {
                    StartEnd.Add("PlannedEndDate", dateEmpty);
                }


                if (ds.Tables[0].Rows[intCount]["RevisedPlanStartDate"].ToString() != "")
                {
                    StartEnd.Add("RevisedPlanStartDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["RevisedPlanStartDate"]));
                }
                else
                {
                    StartEnd.Add("RevisedPlanStartDate", dateEmpty);
                }
                if (ds.Tables[0].Rows[intCount]["RevisedPlanEndDate"].ToString() != "")
                {
                    StartEnd.Add("RevisedPlanEndDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["RevisedPlanEndDate"]));
                }
                else
                {
                    StartEnd.Add("RevisedPlanEndDate", dateEmpty);
                }


                if (ds.Tables[0].Rows[intCount]["ActualStartDate"].ToString() != "")
                {
                    StartEnd.Add("ActualStartDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["ActualStartDate"]));
                }
                else
                {
                    StartEnd.Add("ActualStartDate", dateEmpty);
                }
                if (ds.Tables[0].Rows[intCount]["ActualEndDate"].ToString() != "")
                {
                    StartEnd.Add("ActualEndDate", Convert.ToDateTime(ds.Tables[0].Rows[intCount]["ActualEndDate"]));
                }
                else
                {
                    StartEnd.Add("ActualEndDate", dateEmpty);
                }
                Components.Add(new StagesModel
                {
                    Components = ds.Tables[0].Rows[intCount]["PrimaryStages"].ToString(),
                    ProcessID = ds.Tables[0].Rows[intCount]["SeqID"].ToString() + "-" + ds.Tables[0].Rows[intCount]["Short_Stage"].ToString(),
                    CompletionDays = ds.Tables[0].Rows[intCount]["CompletionDays"].ToString(),
                    othercols = StartEnd,
                    IsEditable= ds.Tables[0].Rows[intCount]["IsEditable"].ToString()

                });
            }

            if (ds.Tables[0].Rows.Count == 0)
            {
                Dictionary<string, DateTime> StartEnd = new Dictionary<string, DateTime>();
                StartEnd.Add("PlannedStartDate", dateEmpty);
                StartEnd.Add("PlannedEndDate", dateEmpty);
                //StartEnd.Add("RevisedPlanStartDate", dateEmpty);
                //StartEnd.Add("RevisedPlanEndDate", dateEmpty);
                //StartEnd.Add("ActualStartDate", dateEmpty);
                //StartEnd.Add("ActualEndDate", dateEmpty);
                Components.Add(new StagesModel
                {
                    CompletionDays = "None",
                    othercols = StartEnd

                });

                ViewBag.CompleChange = "No";
            }
            else { ViewBag.CompleChange = "Yes"; }


            var viewModel = new StagesModel();
            viewModel.ComponentsList = Components;
            return View(viewModel);
            //return View(new List<Models.StagesModel>());
        }
        [SessionExpire]
        public ActionResult GetStagesSchedule()
        {
            try
            {
                DataSet ds = new DataSet();
                string strXML="";
                
                ds = DBProc.GetResultasDataSet("Select Milestone_Stages from " + Session["sCustAcc"].ToString() + "_ProjectManagement where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                strXML = ds.Tables[0].Rows[0]["Milestone_Stages"].ToString().Trim();

                ds = DBProc.GetResultasDataSet("Select Jbm_Autoid,ProcessId,CompletionDays,'' as [PrimaryStages], PlannedStartDate,PlannedEndDate,RevisedPlanStartDate,RevisedPlanEndDate,ActualStartDate,ActualEndDate,Short_Stage from " + Session["sCustAcc"].ToString() + "_ScheduleInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString()); //BK022

                if (strXML != "")
                {
                    XmlDocument objxml = new XmlDocument();
                    objxml.LoadXml(strXML);
                    if (objxml.InnerXml != "Nothing")
                    {

                        for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                        {
                            string strProID = ds.Tables[0].Rows[intCount]["ProcessId"].ToString();
                            string strStg = ds.Tables[0].Rows[intCount]["Short_Stage"].ToString();
                            
                            XmlNodeList xnList = objxml.SelectNodes("//MIS-PID-MilestoneInfo/ProjDet/MilestoneInfo/Stages");
                            foreach (XmlNode xn in xnList)
                            {
                                string pId = xn["ProcessId"].InnerText;
                                string proShortSt = xn["Short_Stage"].InnerText;
                                if (strProID == pId && strStg == proShortSt)
                                {
                                    string processName = xn["Process"].InnerText;
                                    ds.Tables[0].Rows[intCount]["PrimaryStages"] = processName;

                                    break;
                                }
                            }
                        }

                        ds.Tables[0].AcceptChanges();

                    }
                }

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[3].ToString(),a[0].ToString()+"_"+a[1].ToString() + "_" + a[2].ToString(), a[4].ToString(), a[5].ToString(),a[6].ToString(),a[7].ToString(),a[8].ToString(),a[9].ToString(), CreateBtn(a[0].ToString()+"_"+a[1].ToString(), "EditDelete", "Action", "")
                                 };
                // var JSONString = from a in ds.Tables[0].AsEnumerable()
                //                                 select new[] {a[3].ToString(),CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"Input", "CompletionDays",a[2].ToString()), CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "SchStart",a[4].ToString()), CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "SchEnd",a[5].ToString()),CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "RevStart",a[6].ToString()),CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "RevEnd",a[7].ToString()),CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "ActualStart",a[8].ToString()),CreateBtn(a[0].ToString()+"_"+a[1].ToString(),"DateField", "ActualEnd",a[9].ToString()), CreateBtn(a[0].ToString()+"_"+a[1].ToString(), "EditDelete", "Action", "")
                //               };

                return Json(new { dataStg = JSONString }, JsonRequestBehavior.AllowGet);
                //CreateAIBtn(a[7].ToString()
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult AddProjectScheduleData(string sProcessID, string sStageProcessDesc, string sCDays, string sSchStart, string sSchEnd, string sRevStart, string sRevEnd, string sActStart, string sActEnd)
        {
            try
            {
                // To update milestone xml
                DataSet ds = new DataSet();
                //string strXML = "";
                //string shortStage = "";
                //ds = DBProc.GetResultasDataSet("Select Milestone_Stages from " + Session["sCustAcc"].ToString() + "_ProjectManagement where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                //strXML = ds.Tables[0].Rows[0]["Milestone_Stages"].ToString().Trim();
                //if (strXML != "")
                //{
                //    Boolean isNOT = false;
                //    XmlDocument objxml = new XmlDocument();
                //    objxml.LoadXml(strXML);
                //    if (objxml.InnerXml != "Nothing")
                //    {
                //        XmlNodeList xnList = objxml.SelectNodes("//MIS-PID-MilestoneInfo/ProjDet/MilestoneInfo/Stages");

                //        foreach (XmlNode xn in xnList)
                //        {
                //            string pId = xn["ProcessId"].InnerText;
                //            if (sProcessID == pId)
                //            {
                //                isNOT = true;
                //                xn["Status"].InnerText = "1";
                //                shortStage = xn["Short_Stage"].InnerText;
                //                break;
                //            }
                //            else { isNOT = false; }
                //        }
                //    }

                //    if (isNOT == false)
                //    {
                //        // If not exist in DB XML then insert new stage details

                //        strXML = MileStoneXMLNewstage(objxml, sStageProcessDesc, "1");
                //        if (strXML == "Failed")
                //        {
                //            return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                //        }

                //        if (sStageProcessDesc.Trim().IndexOf(" ") > 0)
                //        {
                //            foreach (string av in sStageProcessDesc.Trim().Split(' '))
                //            {
                //                int sCount = av.Length;
                //                if (sCount >= 0 && sCount <= 2)
                //                {
                //                    sCount = 1;
                //                }
                //                else {
                //                    sCount = 2;
                //                }

                //                shortStage += av.Substring(0, sCount);
                //            }
                //        }
                //        else
                //        {
                //            int sCount = sStageProcessDesc.Length;
                //            if (sCount >= 0 && sCount <= 2)
                //            {
                //                sCount = 1;
                //            }
                //            else
                //            {
                //                sCount = 2;
                //            }
                //            shortStage = sStageProcessDesc.Trim().Substring(0, sCount);
                //        }
                //    }

                //    strXML = objxml.OuterXml;
                //    string strStatus = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ProjectManagement SET Milestone_Stages='" + strXML + "' WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                //}
                //else {

                //    string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Project_Config\\MilestoneInfo.xml");

                //    XmlDocument objxml = new XmlDocument();
                //    objxml.Load(strPath);
                //    if (objxml.InnerXml != "Nothing")
                //    {
                //        XmlNodeList xnList = objxml.SelectNodes("//MIS-PID-MilestoneInfo/ProjDet/MilestoneInfo/Stages");
                //        foreach (XmlNode xn in xnList)
                //        {
                //            string pId = xn["ProcessId"].InnerText;
                //            if (sProcessID == pId)
                //            {
                //                xn["Status"].InnerText = "1";
                //                shortStage = xn["Short_Stage"].InnerText;
                //                break;
                //            }
                //        }
                //    }
                //    strXML = objxml.OuterXml;
                //    string strStatus = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ProjectManagement SET Milestone_Stages='" + strXML + "' WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                //}

                //To update Schedule Info

                string dtSchStart, dtSchEnd, dtRevStart, dtRevEnd, dtActStart, dtActEnd;  //If date is empty
                if (string.IsNullOrEmpty(sSchStart)) { dtSchStart = "null"; } else { dtSchStart = "'" + sSchStart + "'"; }
                if (string.IsNullOrEmpty(sSchEnd)) { dtSchEnd = "null"; } else { dtSchEnd = "'" + sSchEnd + "'"; }
                if (string.IsNullOrEmpty(sRevStart)) { dtRevStart = "null"; } else { dtRevStart = "'" + sRevStart + "'"; }
                if (string.IsNullOrEmpty(sRevEnd)) { dtRevEnd = "null"; } else { dtRevEnd = "'" + sRevEnd + "'"; }
                if (string.IsNullOrEmpty(sActStart)) { dtActStart = "null"; } else { dtActStart = "'" + sActStart + "'"; }
                if (string.IsNullOrEmpty(sActEnd)) { dtActEnd = "null"; } else { dtActEnd = "'" + sActEnd + "'"; }
                string strSeqID = sProcessID.Split('|')[0];
                string shortStage = sProcessID.Split('|')[1];

                ds = DBProc.GetResultasDataSet("Select Short_Stage from " + Session["sCustAcc"].ToString() + "_ScheduleInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and Short_Stage='" + shortStage + "' and SeqID=" + strSeqID + "", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count == 0)
                {
                    string strResult = DBProc.GetResultasString("INSERT INTO " + Session["sCustAcc"].ToString() + "_ScheduleInfo (JBM_AutoID,CompletionDays,PlannedStartDate, PlannedEndDate, RevisedPlanStartDate, RevisedPlanEndDate, ActualStartDate, ActualEndDate, Short_Stage, SeqID) VALUES ('" + Session["sJBMAutoID"].ToString() + "','" + sCDays + "'," + dtSchStart + "," + dtSchEnd + "," + dtRevStart + "," + dtRevEnd + "," + dtActStart + "," + dtActEnd + ",'" + shortStage + "','" + strSeqID + "')", Session["sConnSiteDB"].ToString());

                    // To create the StageInfo details
                    string strStatus = gen.CreateStageInfo(Session["sJBMAutoID"].ToString(), Session["sCustAcc"].ToString(), Session["sConnSiteDB"].ToString(), dtSchStart, dtSchEnd, shortStage, "S1004_S1085_S1085");


                }
                else {
                    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ScheduleInfo SET CompletionDays='" + sCDays + "', PlannedStartDate=" + dtSchStart + ", PlannedEndDate=" + dtSchEnd + ", RevisedPlanStartDate=" + dtRevStart + ", RevisedPlanEndDate=" + dtRevEnd + ", ActualStartDate=" + dtActStart + ", ActualEndDate=" + dtActEnd + ", DeleteYN='N' WHERE Short_Stage='" + shortStage + "' and SeqID=" + strSeqID + " and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                }
                
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            };
        }
        string MileStoneXMLNewstage(XmlDocument objxml, string sNewPrimaryStage, string strStatus)
        {
            try
            {
                Dictionary<string, string> collcProcessID = new Dictionary<string, string>();
                Dictionary<string, string> collcShortStg = new Dictionary<string, string>();

                XmlNodeList xnList = objxml.SelectNodes("//MIS-PID-MilestoneInfo/ProjDet/MilestoneInfo/Stages");
                string newProcessID = "";
                foreach (XmlNode xn in xnList)
                {
                    newProcessID = xn["ProcessId"].InnerText;

                    collcProcessID.Add(xn["ProcessId"].InnerText, xn["ProcessId"].InnerText);
                    collcShortStg.Add(xn["Short_Stage"].InnerText, xn["Short_Stage"].InnerText);
                }

                //To generate new process id and short stage
                newProcessID = "PID0" + (Convert.ToInt32(newProcessID.Replace("PID", "").TrimStart(new Char[] { '0' })) + 1).ToString();

                if (collcProcessID.ContainsKey(newProcessID) == true)
                {
                    return "Failed";
                }

                if (sNewPrimaryStage.Trim() != "")
                {

                    //To generate the short stage id
                    string shStage = "";

                    if (sNewPrimaryStage.Trim().IndexOf(" ") > 0)
                    {
                        foreach (string av in sNewPrimaryStage.Trim().Split(' '))
                        {
                            int sCount = av.Length;
                            if (sCount >= 0 && sCount <= 2)
                            {
                                sCount = 1;
                            }
                            else
                            {
                                sCount = 2;
                            }
                            shStage += av.Substring(0, sCount);
                        }
                    }
                    else
                    {
                        int sCount = sNewPrimaryStage.Length;
                        if (sCount >= 0 && sCount <= 2)
                        {
                            sCount = 1;
                        }
                        else
                        {
                            sCount = 2;
                        }
                        shStage = sNewPrimaryStage.Trim().Substring(0, sCount);
                    }

                    if (collcShortStg.ContainsKey(shStage)== true)
                    {
                        return "Failed";
                    }

                    XmlNode stageNode = objxml.SelectSingleNode("//MIS-PID-MilestoneInfo/ProjDet/MilestoneInfo");

                    XmlElement elmNew = objxml.CreateElement("Stages");
                    stageNode.AppendChild(elmNew);

                    elmNew = objxml.CreateElement("Process");
                    elmNew.InnerText = sNewPrimaryStage;
                    stageNode.LastChild.AppendChild(elmNew);

                    elmNew = objxml.CreateElement("ProcessId");
                    elmNew.InnerText = newProcessID;
                    stageNode.LastChild.AppendChild(elmNew);

                    elmNew = objxml.CreateElement("Short_Stage");
                    elmNew.InnerText = shStage;
                    stageNode.LastChild.AppendChild(elmNew);

                    elmNew = objxml.CreateElement("TAT");
                    //elmNew.InnerText = "";
                    stageNode.LastChild.AppendChild(elmNew);

                    elmNew = objxml.CreateElement("Status");
                    elmNew.InnerText = strStatus;
                    stageNode.LastChild.AppendChild(elmNew);

                    stageNode.LastChild.AppendChild(elmNew);
                }
                return objxml.OuterXml.ToString();
            }
             
            catch (Exception)
            {
                return "Failed";
            }
           
        }
        [SessionExpire]
        public ActionResult AddNewPrimaryStage(string sNewPrimaryStage)
        {
            try
            {
                //string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Project_Config\\MilestoneInfo.xml");

                //XmlDocument objxml = new XmlDocument();
                //objxml.Load(strPath);
                //if (objxml.InnerXml != "Nothing")
                //{
                //    string strXML = "";
                //    strXML = MileStoneXMLNewstage(objxml, sNewPrimaryStage, "0");
                //    if (strXML == "Failed")
                //    {
                //        return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                //    }
                //    else
                //    {
                //        objxml.LoadXml(strXML);
                //        objxml.Save(strPath);
                //    }

                //}

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select StageShortName from JBM_StageDescription where StageName='" + sNewPrimaryStage.Trim().Replace("'", "''") + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count == 0)
                {
                    string shortStage = "";
                    if (sNewPrimaryStage.Trim().IndexOf(" ") > 0)
                    {
                        foreach (string av in sNewPrimaryStage.Trim().Split(' '))
                        {
                            int sCount = av.Length;
                            if (sCount >= 0 && sCount <= 2)
                            {
                                sCount = 1;
                            }
                            else
                            {
                                sCount = 2;
                            }

                            shortStage += av.Substring(0, sCount);
                        }
                    }
                    else
                    {
                        int sCount = sNewPrimaryStage.Length;
                        if (sCount >= 0 && sCount <= 2)
                        {
                            sCount = 1;
                        }
                        else
                        {
                            sCount = 2;
                        }
                        shortStage = sNewPrimaryStage.Trim().Substring(0, sCount);
                    }

                    string strResult = DBProc.GetResultasString("INSERT INTO JBM_StageDescription (StageName,StageShortName,StageType,StageTypeId,Is_CustStage,StageSeqID) VALUES ('" + sNewPrimaryStage.Replace("'", "''") + "','" + shortStage + "','Project','2','Y',(Select max(StageSeqID) + 1 from JBM_StageDescription where Is_CustStage='Y'))", Session["sConnSiteDB"].ToString());

                    //clsCollec.WriteLog("New stage " + sNewPrimaryStage + " added", Session["EmpName"].ToString());
                }
                else
                {
                    return Json(new { dataSch = "Exists" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult UpdateStageSchedule(string SaveItemColloc)
        {
            try
            {
                string strCustStageGrp = "";
                if (Session["CustomerName"].ToString() != "TandF")
                {
                    strCustStageGrp = " and CustStageGroup is NULL";
                }
                else
                {
                    if (Regex.IsMatch(Session["StageID"].ToString(), "(PM1|PM2|PM3)", RegexOptions.IgnoreCase))
                        strCustStageGrp = " and CustStageGroup='TF'";
                    else if (Regex.IsMatch(Session["StageID"].ToString(), "(CR)", RegexOptions.IgnoreCase))
                        strCustStageGrp = " and CustStageGroup='TFCR'";
                }

                SaveItemColloc = SaveItemColloc.Replace("01-Jan-1900", "");
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(SaveItemColloc);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string sProcessID = ""; string sCDays = ""; string sSchStart = ""; string sSchEnd = ""; string sRevStart = ""; string sRevEnd = ""; string sActStart = ""; string sActEnd = "";
                        sProcessID = saveIds[i].Split('|')[0];
                        sCDays = saveIds[i].Split('|')[1];
                        sSchStart = saveIds[i].Split('|')[2];
                        sSchEnd = saveIds[i].Split('|')[3];
                        sRevStart = saveIds[i].Split('|')[4];
                        sRevEnd = saveIds[i].Split('|')[5];
                        sActStart = saveIds[i].Split('|')[6];
                        sActEnd = saveIds[i].Split('|')[7];

                        string strSeqID = sProcessID.Split('-')[0];
                        string strShortStg = sProcessID.Split('-')[1];

                        string dtSchStart, dtSchEnd, dtRevStart, dtRevEnd, dtActStart, dtActEnd;
                        if (string.IsNullOrEmpty(sSchStart)) { dtSchStart = "null"; } else { dtSchStart = "'" + sSchStart + "'"; }
                        if (string.IsNullOrEmpty(sSchEnd)) { dtSchEnd = "null"; } else { dtSchEnd = "'" + sSchEnd + "'"; }
                        if (string.IsNullOrEmpty(sRevStart)) { dtRevStart = "null"; } else { dtRevStart = "'" + sRevStart + "'"; }
                        if (string.IsNullOrEmpty(sRevEnd)) { dtRevEnd = "null"; } else { dtRevEnd = "'" + sRevEnd + "'"; }
                        if (string.IsNullOrEmpty(sActStart)) { dtActStart = "null"; } else { dtActStart = "'" + sActStart + "'"; }
                        if (string.IsNullOrEmpty(sActEnd)) { dtActEnd = "null"; } else { dtActEnd = "'" + sActEnd + "'"; }
                        if (string.IsNullOrEmpty(sCDays) || sCDays == "undefined") { sCDays = "null"; } else { sCDays = "'" + sCDays + "'"; }

                        // Update the revised date and actual date as discussed on 02 Jun 2021 with Siva, Shaji, Sujeet
                        string strRevFinstage = "";
                        if (strShortStg.Trim() == "CE")
                        {
                            strRevFinstage = "FP";
                        }
                        else {
                            strRevFinstage =strShortStg.Trim();
                        }
                        int cnt = 0;
                        DataSet ds = new DataSet();
                        ds = DBProc.GetResultasDataSet("Select AutoArtID,DispatchDate,RevisedDate,CeDispDate,CERevisedDate from " + Session["sCustAcc"].ToString() + "_Stageinfo where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = '" + strRevFinstage + "'", Session["sConnSiteDB"].ToString());

                        DataSet ds2 = new DataSet();
                        ds2 = DBProc.GetResultasDataSet("Select CustAcc,CustStageGroup,StageShortID,IsEditable from JBM_StageGroupTAT where StageShortID='" + strRevFinstage + "'" + strCustStageGrp + "", Session["sConnSiteDB"].ToString());

                        if (ds2.Tables[0].Rows[0]["IsEditable"].ToString() == "")  // To update/create primary stages  //Convert.ToInt32(strSeqID) < 14
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                           /// Only update the TAT
                                string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ScheduleInfo SET CompletionDays=" + sCDays + " WHERE Short_Stage='" + strShortStg + "' and SeqID=" + strSeqID + " and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());

                                //*****************Primary stage dates not updated in stageinfo table, shown as min and max date from due and dispatch date**********

                                //for (int k = 0; k < ds.Tables[0].Rows.Count; k++)
                                //{
                                //    string strAutoArtID = ds.Tables[0].Rows[k]["AutoArtID"].ToString();
                                //    string strDispatchDate = ds.Tables[0].Rows[k]["DispatchDate"].ToString();
                                //    string strRevisedDate = ds.Tables[0].Rows[k]["RevisedDate"].ToString();
                                //    string strCeDispDate = ds.Tables[0].Rows[k]["CeDispDate"].ToString();
                                //    string strCERevisedDate = ds.Tables[0].Rows[k]["CERevisedDate"].ToString();

                                //    DataSet ds2 = new DataSet();
                                //    ds2 = DBProc.GetResultasDataSet("Select AutoArtID,DispatchDate,RevisedDate,CeDispDate,CERevisedDate from " + Session["sCustAcc"].ToString() + "_Stageinfo where AutoArtID='" + strAutoArtID + "' and RevFinStage = '" + strRevFinstage + "'", Session["sConnSiteDB"].ToString());
                                //    if (ds2.Tables[0].Rows.Count > 0)
                                //    {
                                //        if (k == 0)
                                //        {
                                //            if (strShortStg.Trim() == "CE")
                                //            {
                                //                string strResult1 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set CeDispDate =" + dtActStart + ",  CERevisedDate=" + dtRevStart + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = 'FP'", Session["sConnSiteDB"].ToString());
                                //            }
                                //            else
                                //            {
                                //                string strResult2 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + dtActStart + ", RevisedDate=" + dtRevStart + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = '" + strShortStg + "'", Session["sConnSiteDB"].ToString());
                                //            }
                                //        }
                                //        else if ((ds.Tables[0].Rows.Count - 1) == cnt)
                                //        {
                                //            if (strShortStg.Trim() == "CE")
                                //            {
                                //                string strResult1 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set CeDispDate =" + dtActEnd + ",  CERevisedDate=" + dtRevEnd + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = 'FP'", Session["sConnSiteDB"].ToString());
                                //            }
                                //            else
                                //            {
                                //                string strResult2 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + dtActEnd + ", RevisedDate=" + dtRevEnd + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = '" + strShortStg + "'", Session["sConnSiteDB"].ToString());
                                //            }
                                //        }
                                //        else
                                //        {
                                //            if (strShortStg.Trim() == "CE")
                                //            {
                                //                string strResult1 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set CeDispDate =" + dtActStart + ",  CERevisedDate=" + dtRevStart + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = 'FP'", Session["sConnSiteDB"].ToString());
                                //            }
                                //            else
                                //            {
                                //                string strResult2 = DBProc.GetResultasString("Update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + dtActStart + ", RevisedDate=" + dtRevStart + "  where AutoArtID ='" + strAutoArtID + "' and RevFinStage = '" + strShortStg + "'", Session["sConnSiteDB"].ToString());
                                //            }
                                //        }
                                //    }
                                //    else
                                //    {
                                //        // To create the StageInfo details
                                //        string strStatus = gen.CreateStageInfo(Session["sJBMAutoID"].ToString(), Session["sCustAcc"].ToString(), Session["sConnSiteDB"].ToString(), dtSchStart, dtSchEnd, strShortStg, "S1004_S1085_S1085");
                                //    }


                                //    cnt = cnt + 1;
                                //}

                            }
                            else
                            {
                                // To create the StageInfo details
                                string strStatus = gen.CreateStageInfo(Session["sJBMAutoID"].ToString(), Session["sCustAcc"].ToString(), Session["sConnSiteDB"].ToString(), dtSchStart, dtSchEnd, strShortStg, "S1004_S1085_S1085");
                            }
                        }
                        else {  // To update other stages

                            //To update Schedule Info
                            string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ScheduleInfo SET CompletionDays=" + sCDays + ", PlannedStartDate=" + dtSchStart + ", PlannedEndDate=" + dtSchEnd + ", RevisedPlanStartDate=" + dtRevStart + ", RevisedPlanEndDate=" + dtRevEnd + ", ActualStartDate=" + dtActStart + ", ActualEndDate=" + dtActEnd + " WHERE Short_Stage='" + strShortStg + "' and SeqID=" + strSeqID + " and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());

                        }
        
                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                //clsCollec.WriteLog("Stage date modified " + SaveItemColloc, Session["EmpName"].ToString());

                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult DeleteStageSchedule(string sProcessID)
        {
            try
            {
                string strSeqID = sProcessID.Split('-')[0];
                string strShortStg = sProcessID.Split('-')[1];

                //To update Schedule Info
                if (Convert.ToInt32(strSeqID) > 13)
                {
                    string strResult = DBProc.GetResultasString("DELETE FROM " + Session["sCustAcc"].ToString() + "_ScheduleInfo WHERE Short_Stage='" + strShortStg + "' and SeqID=" + strSeqID + " and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                }
                else {
                    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ScheduleInfo SET DeleteYN='Y' WHERE Short_Stage='" + strShortStg + "' and SeqID=" + strSeqID + " and JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                }
                
                
                //clsCollec.WriteLog("Stage " + strShortStg + " deleted", Session["EmpName"].ToString());
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult CalculateScheduleDate(string sCompDays, string sStartDate, string CheckedStageID)
        {
            try
            {

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select top 1 PlannedStartDate,SeqID from " + Session["sCustAcc"].ToString() + "_ScheduleInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and SeqID !> 13 order by seqid desc", Session["sConnSiteDB"].ToString()); //BK022

                if (ds.Tables[0].Rows.Count > 0)  // && Convert.ToInt32(ds.Tables[0].Rows[0]["SeqID"].ToString()) < 13
                {
                    sStartDate = ds.Tables[0].Rows[0]["PlannedStartDate"].ToString();
                }

                if (sStartDate == "")
                {
                    return Json(new { dataSch = "" }, JsonRequestBehavior.AllowGet);
                }

                Dictionary<string, string> StartEndCollec = new Dictionary<string, string>();
                


                List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedStageID);
                if (chkIds.Count > 0)
                {
                    for (int i = 0; i < chkIds.Count; i++)
                    {
                        string strStageID = chkIds[i].Split('|')[0];
                        sCompDays = chkIds[i].Split('|')[1];

                        //DateTime dtEndDate;
                        // dtEndDate = AddWorkingDays(Convert.ToDateTime(sStartDate), Convert.ToInt32(sCompDays));
                        DateTime date = Convert.ToDateTime(sStartDate);
                        int daysToAdd = Convert.ToInt32(sCompDays);
                        while (daysToAdd > 0)
                        {
                            date = date.AddDays(1);

                            if (date.DayOfWeek != DayOfWeek.Sunday) //date.DayOfWeek != DayOfWeek.Saturday &&
                            {
                                daysToAdd -= 1;
                            }
                        }

                        StartEndCollec.Add(strStageID, sStartDate + "|" + date.ToString("dd-MMM-yyyy")); //date.ToShortDateString()

                        sStartDate = date.ToShortDateString();  // Re assign for next calculation
                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }
                 
                return Json(new { dataSch = JsonConvert.SerializeObject(StartEndCollec) }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult Components()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult GetComponents()
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select Jbm_Autoid,IntrnlID,ChapterId,NumofMSP,Castoff,StartPage,EndPage,ActualPages,'' as [Deviation],AutoArtID,Active,Seqno from " + Session["sCustAcc"].ToString() + "_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' order by cast(Seqno as int)", Session["sConnSiteDB"].ToString()); //BK022
               
                DataRow[] dr1 = ds.Tables[0].Select("ChapterId ='Front Matter'");
                DataRow[] drr1 = ds.Tables[0].Select("ChapterId ='FM'");
                DataRow[] drr2 = ds.Tables[0].Select("ChapterId ='Frontmatter'");
                if (dr1.Length == 0 && drr1.Length == 0 && drr2.Length==0)
                {
                    BookInModel obj = new BookInModel();
                    obj.ChapterRange = "Front Matter";
                    obj.Platform = "2";
                    obj.StageService = "FP";
                    obj.ChpType = "Chapter";
                    obj.Complexity = "2";
                    insertData(obj);
                }

                DataRow[] dr2 = ds.Tables[0].Select("ChapterId ='Index'");
                if (dr2.Length == 0)
                {
                    BookInModel obj = new BookInModel();
                    obj.ChapterRange = "Index";
                    obj.Platform = "2";
                    obj.StageService = "FP";
                    obj.ChpType = "Chapter";
                    obj.Complexity = "2";
                    insertData(obj);
                }


                DataRow[] dr3 = ds.Tables[0].Select("ChapterId ='Blanks'");
                if (dr3.Length == 0)
                {
                    BookInModel obj = new BookInModel();
                    obj.ChapterRange = "Blanks";
                    obj.Platform = "2";
                    obj.StageService = "FP";
                    obj.ChpType = "Chapter";
                    obj.Complexity = "2";
                    insertData(obj);
                }


                ds = DBProc.GetResultasDataSet("Select Jbm_Autoid,IntrnlID,ChapterId,NumofMSP,Castoff,StartPage,EndPage,ActualPages,'' as [Deviation],AutoArtID,Seqno from " + Session["sCustAcc"].ToString() + "_ChapterInfo WHERE JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active!=0) order by cast(Seqno as int)", Session["sConnSiteDB"].ToString()); //BK022
                dr1 = ds.Tables[0].Select("ChapterId ='Front Matter'");
                drr1 = ds.Tables[0].Select("ChapterId ='FM'");
                drr2 = ds.Tables[0].Select("ChapterId ='Frontmatter'");
                int chaptercount = 0;
                if (dr1.Length != 0 || drr1.Length != 0 || drr2.Length != 0)
                {
                    if (dr1.Length != 0)
                    {
                        DataRow newRow1 = ds.Tables[0].NewRow();
                        newRow1.ItemArray = dr1[0].ItemArray;
                        ds.Tables[0].Rows.Remove(dr1[0]);
                        ds.Tables[0].Rows.InsertAt(newRow1, 0);
                        //chaptercount += 1;
                    }
                    else if (drr1.Length != 0)
                    {
                        DataRow newRow1 = ds.Tables[0].NewRow();
                        newRow1.ItemArray = drr1[0].ItemArray;
                        ds.Tables[0].Rows.Remove(drr1[0]);
                        ds.Tables[0].Rows.InsertAt(newRow1, 0);
                        //chaptercount += 1;
                    }
                    else
                    {
                        DataRow newRow1 = ds.Tables[0].NewRow();
                        newRow1.ItemArray = drr2[0].ItemArray;
                        ds.Tables[0].Rows.Remove(drr2[0]);
                        ds.Tables[0].Rows.InsertAt(newRow1, 0);
                        //chaptercount += 1;
                    }
                }
                dr2 = ds.Tables[0].Select("ChapterId ='Index'");
                if (dr2.Length != 0)
                {
                    DataRow newRow2 = ds.Tables[0].NewRow();
                    newRow2.ItemArray = dr2[0].ItemArray;
                    ds.Tables[0].Rows.Remove(dr2[0]);
                    ds.Tables[0].Rows.InsertAt(newRow2, ds.Tables[0].Rows.Count);
                    chaptercount += 1;
                }
                dr3 = ds.Tables[0].Select("ChapterId ='Blanks'");
                if (dr3.Length != 0)
                {
                    DataRow newRow3 = ds.Tables[0].NewRow();
                    newRow3.ItemArray = dr3[0].ItemArray;
                    ds.Tables[0].Rows.Remove(dr3[0]);
                    ds.Tables[0].Rows.InsertAt(newRow3, ds.Tables[0].Rows.Count);
                    chaptercount += 1;
                }

                //DataRow[] drchp = ds.Tables[0].Select("ChapterId like '%Chp%'");
                //string[] chparray = new string[drchp.Length];
                //if (drchp.Length > 0)
                //{
                //    for (int i = 0; i < drchp.Length; i++)
                //    {
                //        chparray[i] = drchp[i].ItemArray[2].ToString();
                //    }
                //    Array.Sort(chparray, StringComparer.InvariantCulture);
                //}
                //for (int i = 0; i < chparray.Length; i++)
                //{
                //    DataRow[] drch=ds.Tables[0].Select("ChapterId ='"+ chparray[i].ToString()+ "'");
                //    if (drch.Length > 0)
                //    {
                //        DataRow newRow = ds.Tables[0].NewRow();
                //        newRow.ItemArray = drch[0].ItemArray;
                //        ds.Tables[0].Rows.Remove(drch[0]);
                //        ds.Tables[0].Rows.InsertAt(newRow, i);
                //    }
                //}
                int blankid=0;
                int SeqNo = 1;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string ChapterName = ds.Tables[0].Rows[i]["ChapterId"].ToString();
                   
                    if(ChapterName=="Index")
                    {
                        blankid = SeqNo;
                        SeqNo++;                        
                    }
                    if (blankid == 0)
                    {
                        if (ChapterName == "Blanks")
                        {
                            blankid = SeqNo;
                            SeqNo++;
                        }
                       
                    }
                    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Seqno='" + SeqNo + "' WHERE AutoArtID='" + ds.Tables[0].Rows[i]["AutoArtID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                    ds.Tables[0].Rows[i]["Seqno"] = SeqNo;
                    SeqNo++;
                }
                if(blankid == 0)
                    blankid = ds.Tables[0].Rows.Count - chaptercount;
                DataRow newBlankRow1 = ds.Tables[0].NewRow();
                newBlankRow1["Seqno"] = blankid;
                ds.Tables[0].Rows.InsertAt(newBlankRow1, ds.Tables[0].Rows.Count - chaptercount);

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[10].ToString(),
                                     CreateBtn(a[9].ToString(),"Checkbox", "NumofMSP",a[2].ToString()),
                                     (a[2].ToString()!="Blanks"?CreateBtn(a[9].ToString(),"Input", "NumofMSP",a[3].ToString()):""),
                                     CreateBtn(a[9].ToString(),"Input", "Castoff",a[4].ToString()),
                                     (a[2].ToString()=="Front Matter" || a[2].ToString()=="FM" ||a[2].ToString()=="Frontmatter"?CreateBtnn(a[9].ToString(),"Input", "StartPage",a[5].ToString()):CreateBtn(a[9].ToString(),"Input", "StartPage",a[5].ToString())),
                                     (a[2].ToString()=="Front Matter" || a[2].ToString()=="FM" ||a[2].ToString()=="Frontmatter"?CreateBtnn(a[9].ToString(),"Input", "EndPage",a[6].ToString()):CreateBtn(a[9].ToString(),"Input", "EndPage",a[6].ToString())),
                                     ((a[2].ToString()!=""?"<span id='Actual" + a[9].ToString() +  "' data-text='" + CalcActualDeviation("Actual",a[4].ToString(),a[5].ToString(),a[6].ToString(), "No") + "'>" + CalcActualDeviation("Actual",a[4].ToString(),a[5].ToString(),a[6].ToString(), "No") + "</span>":"")),   //a[2].ToString()=="Front Matter" || a[2].ToString()=="FM" ||a[2].ToString()=="Frontmatter"? CreateBtn(a[9].ToString(),"Input", "Actual",a[6].ToString()):
                                     (a[2].ToString()!=""?"<span id='Deviation" + a[9].ToString() +  "' data-text='" + CalcActualDeviation("Deviation",a[4].ToString(),a[5].ToString(),a[6].ToString(), "No") + "'>" + CalcActualDeviation("Deviation",a[4].ToString(),a[5].ToString(),a[6].ToString(), "No") + "</span>":""),
                                     CreateBtn(a[9].ToString(), "EditDelete", "Action", a[2].ToString())
                 };
                return Json(new { dataComp = JSONString ,CCount= chaptercount}, JsonRequestBehavior.AllowGet);
                //CreateAIBtn(a[7].ToString()
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        
        public string CalcActualDeviation(string strCalcType, string strCastoff, string strStartPage, string strEndPage, string strOnChange)
        {
            try
            {
                int intResult = 0;
                int intActual = 0;

                if (Regex.IsMatch(strStartPage, "^[a-zA-Z]*$"))
                {
                    int st = RomanToArabic(strStartPage);
                    strStartPage = st.ToString();
                }



                if (Regex.IsMatch(strEndPage, "^[a-zA-Z]*$"))
                {
                    int ed = RomanToArabic(strEndPage);
                    strEndPage = ed.ToString();
                }


                int intCastOff = 0;
                if (strCastoff == null || strCastoff == "") { intCastOff = 0; } else { intCastOff = Convert.ToInt32(strCastoff); }; ;
                int intStart = 0;
                if (strStartPage == null || strStartPage == "") { intStart = 0; } else { intStart = Convert.ToInt32(strStartPage); }; ;
                int intEnd = 0;
                if (strEndPage == null || strEndPage == "") { intEnd = 0; } else { intEnd = Convert.ToInt32(strEndPage); };

                if (strCalcType == "Actual")
                {
                    //Actual = (End - Start) + 1
                    if (intEnd == 0 && intStart == 0)
                    { intResult = 0;}
                    else {

                        intResult = (intEnd - intStart) + 1;
                        intActual = intResult;
                    }
                }

                if (strCalcType == "Deviation")
                {
                    //Deviation = Actual -  CastOff

                    if (intEnd == 0 && intStart == 0)
                    { intActual = 0; }
                    else { intActual = (intEnd - intStart) + 1; }

                    if (intActual == 0 && intCastOff == 0)
                    { intResult = 0; }
                    else { intResult = intActual - intCastOff; }
                }

                string strFinalResult = "";
                if (strOnChange == "Yes")
                {
                    int intDev = 0;
                    if (strCalcType == "Actual") // To calculate Actual and Deviation
                    {
                        intDev = intActual - intCastOff;
                        strFinalResult = Convert.ToString(intActual + "|" + intDev);
                    }
                    else {
                        strFinalResult = Convert.ToString(intActual + "|" + intResult);
                    }

                }

                else {
                    strFinalResult = Convert.ToString(intResult);
                }

                return strFinalResult;
            }
            catch (Exception ex)
            {
                return "";
            }

        }
        private int RomanToArabic(string roman)
        {
            // Initialize the letter map.
            if (CharValues == null)
            {
                CharValues = new Dictionary<char, int>();
                CharValues.Add('I', 1);
                CharValues.Add('V', 5);
                CharValues.Add('X', 10);
                CharValues.Add('L', 50);
                CharValues.Add('C', 100);
                CharValues.Add('D', 500);
                CharValues.Add('M', 1000);
            }

            if (roman.Length == 0) return 0;
            roman = roman.ToUpper();

            // See if the number begins with (.
            if (roman[0] == '(')
            {
                // Find the closing parenthesis.
                int pos = roman.LastIndexOf(')');

                // Get the value inside the parentheses.
                string part1 = roman.Substring(1, pos - 1);
                string part2 = roman.Substring(pos + 1);
                return 1000 * RomanToArabic(part1) + RomanToArabic(part2);
            }

            // The number doesn't begin with (.
            // Convert the letters' values.
            int total = 0;
            int last_value = 0;
            for (int i = roman.Length - 1; i >= 0; i--)
            {
                int new_value = CharValues[roman[i]];

                // See if we should add or subtract.
                if (new_value < last_value)
                    total -= new_value;
                else
                {
                    total += new_value;
                    last_value = new_value;
                }
            }

            // Return the result.
            return total;
        }
        public string CreateBtn(string uniqueID, string strType, string strColumn, string strField)
        {
            string formControl = string.Empty;
            try  //onkeypress='funcInputValidate()'
            {
                if (strType == "Input")
                {
                    if (strField != "" || uniqueID != "")
                    {
                        if (Session["AccessRights"].ToString() != "Customer")
                            formControl = "<input maxlength='4' onchange=\"funcCalcuationActualDeviation('" + strColumn + "," + uniqueID + "')\"  onkeypress=\"funcInputValidate()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center kpress'  value='" + strField + "'>";
                        else
                            formControl = "<input maxlength='4' onchange=\"funcCalcuationActualDeviation('" + strColumn + "," + uniqueID + "')\"  onkeypress=\"funcInputValidate()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center kpress'  value='" + strField + "' disabled>";

                    }
                    else
                    {
                        if (Session["AccessRights"].ToString() != "Customer")
                            formControl = "<input maxlength='4' onkeypress=\"FuncompoAdd()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "' style='border:1px solid #747779;'>";
                        else
                            formControl = "<input maxlength='4' onkeypress=\"FuncompoAdd()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "' style='border:1px solid #747779;' disabled>";
                    }
                   
                    //if (strColumn == "StartPage" || strColumn == "EndPage")
                    //{
                    //    formControl = "<input maxlength='4' onchange=\"funcCalcuationActualDeviation('" + strColumn + "," + uniqueID + "')\"  onkeypress=\"funcInputValidate()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "'>";
                    //}
                    //else {
                    //    formControl = "<input maxlength='4' onkeypress=\"funcInputValidate()\" type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "'>";
                    //}
                }
                else if (strType == "Checkbox")
                {
                    if (strField != "" || uniqueID != "")
                        formControl = "<input type='checkbox' class='caseChk' id='" + uniqueID + "' onclick=\"funcCheckItem('" + uniqueID + "')\" class='form-control text-center' name='" + uniqueID + "' value='KGLs'>&nbsp;&nbsp;<input type='text' id='txt_" + uniqueID + "' value='" + strField + "' class='form-control' style='width:90%;float:right;'>";
//                    formControl = "<input type='checkbox' class='caseChk' id='" + uniqueID + "' onclick=\"funcCheckItem('" + uniqueID + "')\" class='form-control text-center' name='" + uniqueID + "' value='KGLs'>&nbsp;&nbsp;<label for='" + uniqueID + "'>" + strField + "</label>";
                    else
                        formControl = " <div class='input-group main' style='margin-bottom:25px;'><input style='height:29px;border:1px solid #747779;' type='text' id='chapterrange' value='' ToolTip = 'Range of chapeters ex. 1-10.Single chapter ex. 15.Author names seperated by comma ex.Darwin, Darwin1, Darwin2' required = 'false' data_val = 'false'><label class='control-label' style='color:purple;font-weight:bold;'>Add Components</label> </div>";

                }
                else if (strType == "DateField")
                {
                    if (strField != "")
                    {
                        DateTime dateTime10 = Convert.ToDateTime(strField);
                        strField = dateTime10.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        //strField = DateTime.Now.ToString("yyyy-MM-dd");
                    }

                    if (Session["AccessRights"].ToString() != "Customer")
                        formControl = "<input type='text' id='datepicker1' class='form-control reservation' value='" + strField + "'>";
                    else
                        formControl = "<input type='text' id='datepicker1' class='form-control reservation' value='" + strField + "' disabled>";
                    //formControl = "<input class='form-control reservation' type='date' id='" + strColumn + uniqueID + "'  value='" + strField + "'>";
                }
                else if (strType == "EditDelete")
                {
                    if (strField != "" || uniqueID!="")
                    {
                        if (strColumn == "Action")
                        {
                            if (Session["AccessRights"].ToString() != "Customer")
                                //formControl = "<table><tr><td><a id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\" class='btn' href='javascript:void();' style='color:#838e83' tooltip='Edit'><i class='fa fa-edit' ></i></a></td><td><a id='Update" + uniqueID + "' onClick=\"compoUpdate('" + uniqueID + "')\" class='btn' href='javascript:void();' style='color:#6610f2;display:none;' tooltip='Update'><i class='fa fa-save' ></i></a></td><td><a id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\"  style='color:#eb2227' class='btn' href='javascript:void();'  tooltip='Delete'><i class='fa fa-trash' ></i></a></td></tr></table>";
                                formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\" style='display:none;' class='btn btn-light' name='edit' value='Edit'><span class='fas fa-edit fa-1x text-grey'></span></button><button type='button' id='Update" + uniqueID + "' onClick=\"compoUpdate('" + uniqueID + "')\" class='btn btn-light HideItem' name='update' value='Update'><span class='fas fa-save fa-1x text-blue'></span></button><button type='button' id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\"  class='btn btn-light HideItem' name='delete' value='Delete'><span class='fas fa-trash fa-1x text-red'></span></button>";
                            else
                                formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\" style='display:none;' class='btn btn-light' name='edit' value='Edit'><span class='fas fa-edit fa-1x text-grey'></span></button><button type='button' id='Update" + uniqueID + "' onClick=\"compoUpdate('" + uniqueID + "')\" style='display:none;' class='btn btn-light HideItem' name='update' value='Update'><span class='fas fa-save fa-1x text-blue'></span></button><button type='button' id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\" style='display:none;' class='btn btn-light HideItem' name='delete' value='Delete'><span class='fas fa-trash fa-1x text-red'></span></button>";

                        }
                    }
                    else {
                        if (Session["AccessRights"].ToString() != "Customer")
                            //formControl = "<table><tr><td><a id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\" class='btn' href='javascript:void();' style='color:#838e83' tooltip='Edit'><i class='fa fa-edit' ></i></a></td><td><a id='Update" + uniqueID + "' onClick=\"compoUpdate('" + uniqueID + "')\" class='btn' href='javascript:void();' style='color:#6610f2;display:none;' tooltip='Update'><i class='fa fa-save' ></i></a></td><td><a id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\"  style='color:#eb2227' class='btn' href='javascript:void();'  tooltip='Delete'><i class='fa fa-trash' ></i></a></td></tr></table>";
                            formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoAdd('" + uniqueID + "')\"  class='btn btn-light' name='edit' value='Edit'><span class='fas fa-plus-square fa-1x text-grey'></span></button>";
                        else
                            formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoAdd('" + uniqueID + "')\"  class='btn btn-light' name='edit' value='Edit' style='display:none;'><span class='fas fa-plus-square fa-1x text-grey'></span></button>";

                    }

                }


                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        public string CreateBtnn(string uniqueID, string strType, string strColumn, string strField)
        {
            string formControl = string.Empty;
            try  
            {
                if (strType == "Input")
                {
                    if (Session["AccessRights"].ToString() != "Customer")
                        formControl = "<input maxlength='12' onchange=\"funcCalcuationActualDeviation('" + strColumn + "," + uniqueID + "')\"   type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "'>";
                    else
                        formControl = "<input maxlength='12' onchange=\"funcCalcuationActualDeviation('" + strColumn + "," + uniqueID + "')\"   type='text' id='" + strColumn + uniqueID + "' class='form-control text-center'  value='" + strField + "' disabled>";
                    
                }

                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult ComponentsUpdate(string SaveItemColloc) //string sAutoArtID, string sNumofMSP, string sCastoff, string sStartPage, string sEndPage
        {

            try
            {
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(SaveItemColloc);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string sChapterID = "";string sAutoArtID = ""; string sNumofMSP = ""; string sCastoff = ""; string sStartPage = ""; string sEndPage = ""; string sActualPage = "";
                        sAutoArtID = saveIds[i].Split('|')[0];
                        sNumofMSP = saveIds[i].Split('|')[1];
                        sCastoff = saveIds[i].Split('|')[2];
                        sStartPage = saveIds[i].Split('|')[3];
                        sEndPage = saveIds[i].Split('|')[4];
                        sActualPage = saveIds[i].Split('|')[5];
                        sChapterID= saveIds[i].Split('|')[6];
                        if (sNumofMSP == "undefined")
                        {
                            sNumofMSP = "";
                        }
                        //To update 
                        string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET NumofMSP='" + sNumofMSP + "', Castoff='" + sCastoff + "', StartPage='" + sStartPage + "', EndPage='" + sEndPage + "', ActualPages='" + sActualPage + "',ChapterID='"+ sChapterID + "' WHERE AutoArtID='" + sAutoArtID + "'", Session["sConnSiteDB"].ToString());

                    }

                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

            //try
            //{
            //    //To update 
            //    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET NumofMSP='" + sNumofMSP + "', Castoff='" + sCastoff + "', StartPage='" + sStartPage + "', EndPage='" + sEndPage + "' WHERE AutoArtID='" + sAutoArtID + "'", Session["sConnSiteDB"].ToString());

            //    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            //}
            //catch (Exception)
            //{
            //    return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            //}
        }

        [SessionExpire]
        public ActionResult ComponentsSeqnoUpdate(string SaveItemColloc) 
        {

            try
            {
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(SaveItemColloc);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string sSeqno = ""; string sAutoArtID = ""; 
                        sAutoArtID = saveIds[i].Split('|')[0];
                        sSeqno = saveIds[i].Split('|')[1];
                        //To update 
                        if (sAutoArtID != "chapterrange")
                        {
                            string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Seqno='" + sSeqno + "' WHERE AutoArtID='" + sAutoArtID + "'", Session["sConnSiteDB"].ToString());
                        }
                    }

                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult ComponentsDelete(string sAutoArtID)
        {
            try
            {
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(sAutoArtID);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string ssAutoArtID = ""; 
                        ssAutoArtID = saveIds[i];

                        DataTable dt= DBProc.GetResultasDataTbl(" select ChapterId from BK_ChapterInfo where AutoArtID='"+ ssAutoArtID + "'", Session["sConnSiteDB"].ToString());
                        if(dt.Rows.Count>0)
                        {
                            string ChapterID = dt.Rows[0]["ChapterId"].ToString();
                            if(ChapterID=="Front Matter" || ChapterID == "Frontmatter" || ChapterID == "FM" || ChapterID == "Index" || ChapterID == "Blanks")
                            {
                                //To Update
                                string strupdateResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Active='0',NumofMSP=null,Castoff=null,StartPage=null,EndPage=null,ActualPages=null WHERE AutoArtID='" + ssAutoArtID + "'", Session["sConnSiteDB"].ToString());
                            }
                            else
                            {
                                //To Delete 
                                string strResult = DBProc.GetResultasString("DELETE FROM " + Session["sCustAcc"].ToString() + "_ProdInfo WHERE AutoArtID='" + ssAutoArtID + "'", Session["sConnSiteDB"].ToString());
                                string strResult1 = DBProc.GetResultasString("DELETE FROM " + Session["sCustAcc"].ToString() + "_ChapterInfo WHERE AutoArtID='" + ssAutoArtID + "'", Session["sConnSiteDB"].ToString());
                           }
                        }
                         }

                    return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
                }
                
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }      

        public JsonResult AddNewComponents(string schpRange,string sNumofMSP,string sCastoff,string sStartPage,string sEndPage)
        {
            try
            {
                // If chapters entered as range page no not updated only for single chapters
                Regex Regex2 = new Regex(@"([0-9]\-[0-9])", RegexOptions.IgnoreCase);
                Match m2 = Regex2.Match(schpRange.ToString());
                if (m2.Success)
                {
                    sNumofMSP = "";
                }

                string SeqNo = "";
                DataTable dtmaxid = DBProc.GetResultasDataTbl(" Select max(CAST(Seqno AS INT)) from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                if (dtmaxid.Rows.Count > 0)
                {
                    if (dtmaxid.Rows[0][0].ToString().Trim() != "")
                        SeqNo = (Convert.ToInt32(dtmaxid.Rows[0][0].ToString()) + 1).ToString();
                    else
                        SeqNo = "1";
                }
                else
                {
                    SeqNo = "1";
                }

                BookInModel obj = new BookInModel();
                obj.ChapterRange = schpRange;
                obj.Platform = "2";
                obj.StageService = "FP";
                obj.ChpType = "Chapter";
                obj.Complexity = "2";
                if (obj.ChapterRange == null)
                {
                    return Json(new { dataSch = "Range" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    string ChapterID = obj.ChapterRange;
                    if (ChapterID == "Front Matter" || ChapterID == "Frontmatter" || ChapterID == "FM" || ChapterID == "Index" || ChapterID == "Blanks")
                    {
                        DataTable dt = DBProc.GetResultasDataTbl(" select Active from BK_ChapterInfo where ChapterId='" + ChapterID + "' and Jbm_Autoid='"+ Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                        if (dt.Rows.Count > 0)
                        {
                            string strActive = dt.Rows[0]["Active"].ToString();
                            if(strActive=="0")
                            { 
                                //To Update
                                string strupdateResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Active='1' WHERE ChapterId='" + ChapterID + "' and Jbm_Autoid='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                            }
                        }
                        else
                        {
                            insertData(obj);
                            string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Seqno='" + SeqNo + "', NumofMSP='" + sNumofMSP + "' WHERE Seqno is null and Jbm_Autoid='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                        }
                    }
                    else
                    {
                        insertData(obj);
                        string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_ChapterInfo SET Seqno='" + SeqNo + "',NumofMSP='" + sNumofMSP + "' WHERE Seqno is null and Jbm_Autoid='" + Session["sJBMAutoID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                    }

                }

                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult Schedule()
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            List<ScheduleModel> Components = new List<ScheduleModel>();
            List<ScheduleModel> Schedules = new List<ScheduleModel>();
            con.Open();
            SqlCommand com = new SqlCommand("select AutoArtId, ChapterID from " + Session["sCustAcc"].ToString() + "_chapterinfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and (Active is null or Active!=0) order by cast(Seqno as int)", con);
            SqlDataAdapter da = new SqlDataAdapter(com);
            string cmdstring = "";
            string strCustStageGrp = string.Empty;

            if (Session["CustomerName"].ToString() != "TandF")
            {
                strCustStageGrp = " and CustStageGroup is NULL";
            }
            else
            {
                if (Regex.IsMatch(Session["StageID"].ToString(), "(PM1|PM2|PM3)", RegexOptions.IgnoreCase))
                    strCustStageGrp = " and CustStageGroup='TF'";
                else if (Regex.IsMatch(Session["StageID"].ToString(), "(CR)", RegexOptions.IgnoreCase))
                    strCustStageGrp = " and CustStageGroup='TFCR'";
            }
            cmdstring = "select  distinct SI.ProcessId,SI.Short_Stage,SI.CompletionDays,SI.SeqID from " + Session["sCustAcc"].ToString() + "_scheduleinfo SI join JBM_StageDescription SD join JBM_StageGroupTAT SG  on SD.StageShortName=SG.StageShortID on SD.StageShortName=SI.Short_Stage where SI.JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "' and SG.IsEditable is null and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + " and SD.IS_CustStage = 'Y' and  (SI.DeleteYN ='N' or SI.DeleteYN is null) order by SI.SeqID asc";

            SqlCommand cmdProcessID = new SqlCommand(cmdstring, con);
            SqlDataAdapter daProcessID = new SqlDataAdapter(cmdProcessID);
            //SqlCommand cmdxml = new SqlCommand("select milestone_stages from " + Session["sCustAcc"].ToString() + "_projectmanagement where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'", con);
            //string xmlfile = Convert.ToString(cmdxml.ExecuteScalar());
            SqlCommand cmdRevFinStage = new SqlCommand("select distinct RevFinStage from " + Session["sCustAcc"].ToString() + "_stageinfo where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo where JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "')", con);
            SqlDataAdapter daRevFinStage = new SqlDataAdapter(cmdRevFinStage);

            DataTable dtProcessID = new DataTable();
            daProcessID.Fill(dtProcessID);
            string cmdstring1 = "";
            //if (Session["CustomerName"].ToString() != "TandF")
            //else
            //    cmdstring1 = "Select distinct SD.StageName,SI.Short_Stage as [StageShortName],SI.SeqID as [StageSeqID] from " + Session["sCustAcc"].ToString() + "_scheduleinfo SI join JBM_StageDescription SD on SD.StageShortName=SI.Short_Stage where SI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and SD.IS_CustStage = 'Y' and SD.CustType='TandF' and SD.CustStageID like '" + Session["StageID"] + "' and  (SI.DeleteYN ='N' or SI.DeleteYN is null)  and SD.StageSeqID < 14 order by SI.SeqID asc";
            cmdstring1 = "Select distinct SD.StageName,SI.Short_Stage as [StageShortName],SI.SeqID as [StageSeqID] from " + Session["sCustAcc"].ToString() + "_scheduleinfo SI join JBM_StageDescription SD join JBM_StageGroupTAT SG  on SD.StageShortName=SG.StageShortID on SD.StageShortName=SI.Short_Stage where SI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and SD.IS_CustStage = 'Y' and CustAcc='" + Session["CustomerName"].ToString() + "'" + strCustStageGrp + " and  (SI.DeleteYN ='N' or SI.DeleteYN is null)  and SG.IsEditable is null order by SI.SeqID asc";

            DataSet dschedules = new DataSet();

            //dschedules = DBProc.GetResultasDataSet("Select distinct SD.StageName,SD.StageShortName,SD.StageSeqID from JBM_StageDescription SD join  " + Session["sCustAcc"].ToString() + "_scheduleinfo SI on SD.StageShortName=SI.Short_Stage where SI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and SD.IS_CustStage = 'Y'  and SD.StageSeqID < 14 order by SD.StageSeqID asc", Session["sConnSiteDB"].ToString());
            dschedules = DBProc.GetResultasDataSet(cmdstring1, Session["sConnSiteDB"].ToString());

            for (int intCount = 0; intCount < dschedules.Tables[0].Rows.Count; intCount++)
            {
                string strStageName = dschedules.Tables[0].Rows[intCount]["StageName"].ToString();
                string strShortStg = dschedules.Tables[0].Rows[intCount]["StageShortName"].ToString();
                string strStgSeq = dschedules.Tables[0].Rows[intCount]["StageSeqID"].ToString();
                Schedules.Add(new ScheduleModel
                {
                    Schedules = Convert.ToString(strStageName)
                });
            }

            DataTable dt = new DataTable();
            da.Fill(dt);
            DataTable dtRevFinStage = new DataTable();
            daRevFinStage.Fill(dtRevFinStage);



            foreach (DataRow dr in dt.Rows)
            {
                Dictionary<string, DateTime?> othercol = new Dictionary<string, DateTime?>();
                List<string> schedulecol = new List<string>();
                List<string> CompletionDay = new List<string>();
                foreach (DataRow drProcessID in dtProcessID.Rows)
                {
                    SqlCommand cmddatedata = new SqlCommand();
                    //if (drProcessID["Short_Stage"].ToString().Trim() == "CE")
                    //{
                    //    cmddatedata = new SqlCommand("SELECT  SI.CeDispDate as ReceivedDate, SI.CeDueDate as DueDate,CERevisedDate as RevisedDate FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI INNER JOIN  " + Session["sCustAcc"].ToString() + "_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID where CI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and   SI.RevFinStage = 'FP' and CI.ChapterID = '" + Convert.ToString(dr["ChapterID"]) + "'", con);
                    //}
                    //else
                    //{
                    //    cmddatedata = new SqlCommand("SELECT  SI.DispatchDate as ReceivedDate, SI.DueDate,SI.RevisedDate FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI INNER JOIN  " + Session["sCustAcc"].ToString() + "_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID where CI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and   SI.RevFinStage = '" + Convert.ToString(drProcessID["Short_Stage"]) + "' and CI.ChapterID = '" + Convert.ToString(dr["ChapterID"]) + "'", con);
                    //}
                    if (drProcessID["Short_Stage"].ToString().Trim() == "CE")
                    {
                        cmddatedata = new SqlCommand("SELECT  FORMAT (CAST(SI.CeDispDate as date), 'dd/MMM/yyyy') as ReceivedDate, FORMAT (CAST(SI.CeDueDate as date), 'dd/MMM/yyyy') as DueDate,FORMAT (CAST(CERevisedDate as date), 'dd/MMM/yyyy') as RevisedDate FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI INNER JOIN  " + Session["sCustAcc"].ToString() + "_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID where CI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and   SI.RevFinStage = 'FP' and CI.ChapterID = '" + Convert.ToString(dr["ChapterID"]) + "'", con);
                    }
                    else if (drProcessID["Short_Stage"].ToString().Trim() == "1pco")
                    {
                        cmddatedata = new SqlCommand("SELECT  FORMAT (CAST(SI.PrintFinalDue as date), 'dd/MMM/yyyy') as Pub_Date,FORMAT (CAST(SI.PR_corr_Appr as date), 'dd/MMM/yyyy') as PR_Date,FORMAT (CAST(SI.Aut_Corr_Appr as date), 'dd/MMM/yyyy') as AU_Date, FORMAT (CAST(SI.DueDate as date) , 'dd/MMM/yyyy') as DueDate,FORMAT (CAST(SI.RevisedDate as date), 'dd/MMM/yyyy') as RevisedDate FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI INNER JOIN  " + Session["sCustAcc"].ToString() + "_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID where CI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and   SI.RevFinStage = '" + Convert.ToString(drProcessID["Short_Stage"]) + "' and CI.ChapterID = '" + Convert.ToString(dr["ChapterID"]) + "'", con);
                    }
                    else
                    {
                        cmddatedata = new SqlCommand("SELECT  FORMAT (CAST(SI.DispatchDate as date), 'dd/MMM/yyyy') as ReceivedDate, FORMAT (CAST(SI.DueDate as date) , 'dd/MMM/yyyy') as DueDate,FORMAT (CAST(SI.RevisedDate as date), 'dd/MMM/yyyy') as RevisedDate FROM  " + Session["sCustAcc"].ToString() + "_ChapterInfo CI INNER JOIN  " + Session["sCustAcc"].ToString() + "_Stageinfo SI ON CI.AutoArtID = SI.AutoArtID where CI.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and   SI.RevFinStage = '" + Convert.ToString(drProcessID["Short_Stage"]) + "' and CI.ChapterID = '" + Convert.ToString(dr["ChapterID"]) + "'", con);
                    }
                    SqlDataAdapter dadatedata = new SqlDataAdapter(cmddatedata);
                    DataTable dtdatedata = new DataTable();
                    dadatedata.Fill(dtdatedata);
                    schedulecol.Add(Convert.ToString(drProcessID["Short_Stage"]));
                    CompletionDay.Add(Convert.ToString(drProcessID["CompletionDays"]));
                    if (dtdatedata.Rows.Count > 0)
                    {
                        if (Convert.ToString(drProcessID["Short_Stage"]) == "1pco")
                        {
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_DueDate", Convert.ToString(dtdatedata.Rows[0]["DueDate"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["DueDate"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_RevisedDate", Convert.ToString(dtdatedata.Rows[0]["RevisedDate"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["RevisedDate"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_PubDate", Convert.ToString(dtdatedata.Rows[0]["Pub_Date"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["Pub_Date"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_PRDate", Convert.ToString(dtdatedata.Rows[0]["PR_Date"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["PR_Date"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_AUDate", Convert.ToString(dtdatedata.Rows[0]["AU_Date"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["AU_Date"]) : (DateTime?)null);
                        }
                        else {
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_DueDate", Convert.ToString(dtdatedata.Rows[0]["DueDate"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["DueDate"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_RevisedDate", Convert.ToString(dtdatedata.Rows[0]["RevisedDate"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["RevisedDate"]) : (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_ReceivedDate", Convert.ToString(dtdatedata.Rows[0]["ReceivedDate"]) != "" ? Convert.ToDateTime(dtdatedata.Rows[0]["ReceivedDate"]) : (DateTime?)null);
                        }
                        

                    }
                    else
                    {

                        if (Convert.ToString(drProcessID["Short_Stage"]) == "1pco")
                        {
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_DueDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_RevisedDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_PubDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_PRDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_AUDate", (DateTime?)null);
                        }
                        else
                        {
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_DueDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_RevisedDate", (DateTime?)null);
                            othercol.Add(Convert.ToString(drProcessID["Short_Stage"]) + "_ReceivedDate", (DateTime?)null);
                        }
                      
                    }
                }
                Components.Add(new ScheduleModel
                {
                    ProcessID = Convert.ToString(dr["AutoArtID"]), //Convert.ToString(dr["ChapterID"]).Replace(" ", "")
                    Components = Convert.ToString(dr["ChapterID"]),
                    othercols = othercol,
                    schedulecols = schedulecol,
                    CompletionDays = CompletionDay
                });

            }
            con.Close();

            var rowfm1 = Components.Where(c => c.Components == "Front Matter");
            var rowfm2 = Components.Where(c => c.Components == "Frontmatter");
            var rowfm3 = Components.Where(c => c.Components == "FM");
            var row1 = rowfm1.ToList();
            var row2 = rowfm2.ToList();
            var row3 = rowfm3.ToList();

            if (row1.Count != 0 || row2.Count != 0 || row3.Count != 0)
            {
                if (row1.Count != 0)
                {
                    Components.RemoveAll(x => x.Components == "Front Matter");
                    Components.InsertRange(0, row1.ToList());
                }
                else if (row2.Count != 0)
                {
                    Components.RemoveAll(x => x.Components == "Frontmatter");
                    Components.InsertRange(0, row2.ToList());
                }
                else
                {
                    Components.RemoveAll(x => x.Components == "FM");
                    Components.InsertRange(0, row3.ToList());
                }
            }

            var rowindex = Components.Where(c => c.Components == "Index");
            var rowIn = rowindex.ToList();
            if (rowIn.Count != 0)
            {
                Components.RemoveAll(x => x.Components == "Index");
                Components.InsertRange(Components.Count, rowIn.ToList());
            }


            var rowBlanks = Components.Where(c => c.Components == "Blanks");
            var rowBla = rowBlanks.ToList();
            if (rowBla.Count != 0)
            {
                Components.RemoveAll(x => x.Components == "Blanks");
                //Components.InsertRange(Components.Count, rowBla.ToList());
            }
            //var drchp = Components.Where(c => c.Components.Contains("Chp"));
            //var rowdrchp = drchp.ToList();
            //string[] chparray = new string[rowdrchp.Count];
            //if (rowdrchp.Count > 0)
            //{
            //    for (int i = 0; i < rowdrchp.Count; i++)
            //    {
            //        chparray[i] = rowdrchp[i].Components.ToString();
            //    }
            //    Array.Sort(chparray, StringComparer.InvariantCulture);
            //}
            //for (int i = 0; i < chparray.Length; i++)
            //{
            //    var drch = Components.Where(c => c.Components == "" + chparray[i].ToString() + "");
            //    var rowdrch = drch.ToList();
            //    if (rowdrch.Count > 0)
            //    {
            //        Components.RemoveAll(x => x.Components == "" + chparray[i].ToString() + "");
            //        Components.InsertRange(i, rowdrch.ToList());
            //    }
            //}

            var viewModel = new ScheduleListModel();
            viewModel.ScheduleList = Schedules;
            viewModel.ComponentsList = Components;

            return View(viewModel);
        }
        [HttpPost]
        [SessionExpire]
        public ActionResult SaveSchedule(string vSchedulearray)
        {
            try
            {
                string strLogColl = "";

                vSchedulearray = vSchedulearray.Replace("01-Jan-1900", "");

                var result = JsonConvert.DeserializeObject<List<vSchedule>>(vSchedulearray);
                SqlConnection con = new SqlConnection();
                foreach (var schedulearray in result)
                {
                    //Console.WriteLine(schedulearray.UserID);
                    string SChapterID = schedulearray.SChapterID.ToString();
                    string SShort_Stage = schedulearray.SShort_Stage.ToString();

                    string SDueDate = schedulearray.SDueDate != "" ? "'" + schedulearray.SDueDate + "'" : "null";
                    string SReceivedDate = schedulearray.SReceivedDate != "" ? "'" + schedulearray.SReceivedDate + "'" : "null";
                    string SRevisedDate = schedulearray.SRevisedDate != "" ? "'" + schedulearray.SRevisedDate + "'" : "null";

                    string SPubDate = schedulearray.SPubDate != "" ? "'" + schedulearray.SPubDate + "'" : "null";
                    string SPRDate = schedulearray.SPRDate != "" ? "'" + schedulearray.SPRDate + "'" : "null";
                    string SAUDate = schedulearray.SAUDate != "" ? "'" + schedulearray.SAUDate + "'" : "null";

                    // To create the stage if stages is missing
                    string shortStage = string.Empty;
                    if (SShort_Stage == "CE")
                    {
                        shortStage = "FP";
                    }
                    else { shortStage = SShort_Stage; }

                    DataSet ds1 = new DataSet();
                    ds1 = DBProc.GetResultasDataSet("Select AutoArtID from " + Session["sCustAcc"].ToString() + "_StageInfo where AutoArtID='" + SChapterID + "' and RevFinStage='" + shortStage + "'", Session["sConnSiteDB"].ToString());
                    
                    if (ds1.Tables[0].Rows.Count == 0)
                    {
                        string strStatus = DBProc.GetResultasString("INSERT INTO " + Session["sCustAcc"].ToString() + "_StageInfo (AutoArtID,RevFinStage,ArtStageTypeID, Rev_wf,ReceivedDate,DueDate) VALUES ('" + SChapterID + "','" + SShort_Stage + "','S1004_S1085_S1085_40385_" + DateTime.Now.ToString("dd-MM-yyyy h:mm tt") + "','W143',CONVERT(DateTime, " + SReceivedDate + ", 101),CONVERT(DateTime, " + SDueDate + ", 101))", Session["sConnSiteDB"].ToString());
                    }
                    ///end stage creation

                    con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                    con.Open();
                    if (SShort_Stage.Trim() == "CE")
                    {
                        SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set CeDispDate =" + SReceivedDate + ", CeDueDate=" + SDueDate + ", CERevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where AutoArtID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = 'FP'", con);
                        cmd.ExecuteNonQuery();
                        strLogColl = strLogColl + Environment.NewLine + "AutoArtID:" + SChapterID + ": CeDispDate = " + SReceivedDate + ", CeDueDate = " + SDueDate + ", CERevisedDate = " + SRevisedDate;
                    }
                    else if (SShort_Stage.Trim() == "1pco")
                    {
                        SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set PrintFinalDue =" + SPubDate + ",PR_corr_Appr =" + SPRDate + ",Aut_Corr_Appr =" + SAUDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where AutoArtID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = '" + SShort_Stage + "'", con);
                        cmd.ExecuteNonQuery();
                        strLogColl = strLogColl + Environment.NewLine + "AutoArtID:" + SChapterID + ": PrintFinalDue =" + SPubDate + ",PR_corr_Appr =" + SPRDate + ",Aut_Corr_Appr =" + SAUDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate;
                    }
                    else
                    {
                        //SqlCommand cmd = new SqlCommand(" update BK_Stageinfo set ReceivedDate =null, DueDate=null,RevisedDate=null  where AutoArtID in (select AutoArtID from BK_ChapterInfo  where ChapterID = 'Chp01' and JBM_AutoID = 'BK510') and RevFinStage = 'FP'", con);
                        SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + SReceivedDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where AutoArtID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = '" + SShort_Stage + "'", con);
                        cmd.ExecuteNonQuery();
                        strLogColl = strLogColl + Environment.NewLine + "AutoArtID:" + SChapterID + ": DispatchDate =" + SReceivedDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate;
                    }

                    con.Close();

                    
                }

                //SqlConnection con = new SqlConnection();
                //int count = schedulearray.Count;
                //for (int i = 0; i < count; i++)
                //{
                //   // string SChapterID = schedulearray[i].SChapterID.ToString();
                //   // string SShort_Stage = schedulearray[i].SShort_Stage.ToString();

                //    //string SDueDate = schedulearray[i].SDueDate != null ? "'" + schedulearray[i].SDueDate + "'" : "null";
                //    //string SReceivedDate = schedulearray[i].SReceivedDate != null ? "'" + schedulearray[i].SReceivedDate + "'" : "null";
                //    //string SRevisedDate = schedulearray[i].SRevisedDate != null ? "'" + schedulearray[i].SRevisedDate + "'" : "null";
                //    //con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                //    //con.Open();
                //    //if (SShort_Stage.Trim() == "CE")
                //    //{
                //    //    SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set CeDispDate =" + SReceivedDate + ", CeDueDate=" + SDueDate + ", CERevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where AutoArtID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = 'FP'", con);
                //    //    cmd.ExecuteNonQuery();
                //    //}
                //    //else
                //    //{
                //    //    //SqlCommand cmd = new SqlCommand(" update BK_Stageinfo set ReceivedDate =null, DueDate=null,RevisedDate=null  where AutoArtID in (select AutoArtID from BK_ChapterInfo  where ChapterID = 'Chp01' and JBM_AutoID = 'BK510') and RevFinStage = 'FP'", con);
                //    //    SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + SReceivedDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where AutoArtID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = '" + SShort_Stage + "'", con);
                //    //    cmd.ExecuteNonQuery();
                //    //}
                //    //con.Close();
                //}
                //clsCollec.WriteLog("Schedule Tab " + strLogColl + " updated", Session["EmpName"].ToString());
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult CalcScheduleDate(string sCompDays, string sStartDate, string CheckedStageID, string chkChapter, string vExistDueDate)
        {
            try
            {
                //if (sPreviousDate == "")
                //{
                //    sPreviousDate = sStartDate;
                //}
                //DateTime datePrevious = Convert.ToDateTime(sPreviousDate);
                //TimeSpan nod = (Convert.ToDateTime(sStartDate) - datePrevious);
                //int incrementDays = nod.Days;

                Dictionary<string, string> StartEndCollec = new Dictionary<string, string>();
                List<string> collExistDueDate = JsonConvert.DeserializeObject<List<string>>(vExistDueDate);
                Dictionary<string, string> dicExistsStart = new Dictionary<string, string>();
                if (collExistDueDate.Count > 0)
                {

                    string pickedStage = CheckedStageID.Split('|')[2].ToString();
                    bool blnIsPicked = false;
                    for (int m = 0; m < collExistDueDate.Count; m++)
                    {
                        string strAutoArtID = collExistDueDate[m].Split('|')[0];
                        string strSStage = collExistDueDate[m].Split('|')[1];
                        string strSDate = collExistDueDate[m].Split('|')[2];
                        if (pickedStage == strSStage && blnIsPicked == false)
                        {
                            blnIsPicked = true;
                            dicExistsStart.Add(strAutoArtID + strSStage, sStartDate);
                        }
                        if (strSDate == "undefined")
                        {
                            return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                        }

                        if (strSDate != "" && strSDate != "undefined")
                        {
                            if (dicExistsStart.ContainsKey(strAutoArtID + strSStage) == false)
                            {
                                dicExistsStart.Add(strAutoArtID + strSStage, strSDate);
                            }

                        }
                    }
                }



                List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedStageID);
                List<string> chkchapterarrayIds = JsonConvert.DeserializeObject<List<string>>(chkChapter);
                string[] chkchapterarray = chkchapterarrayIds.ToArray();
                if (chkIds.Count > 0)
                {
                    for (int i = 0; i < chkIds.Count; i++)
                    {
                        string strChapterID = chkIds[i].Split('|')[0];
                        sCompDays = chkIds[i].Split('|')[1];
                        string short_stage = chkIds[i].Split('|')[2];
                        string PickedDatename = chkIds[i].Split('|')[3];
                        string Short_Stagelst = chkIds[i].Split('|')[4];
                        string CompeltionDays = chkIds[i].Split('|')[5];
                        string[] Short_Stagelstarray = Short_Stagelst.Split(',');
                        string[] CompeltionDaysarray = CompeltionDays.Split(',');
                        int flag = 0;
                        
                        for (int k = 0; k < Short_Stagelstarray.Length; k++)
                        {
                           
                            if (Short_Stagelstarray[k] == short_stage)
                            {
                                flag = 1;
                                if (flag == 1)
                                {

                                    for (int l = 0; l < chkchapterarray.Length; l++)
                                    {
                                        int flagstart = 0;
                                        string nstartdate = sStartDate;
                                       
                                        for (int j = k; j < Short_Stagelstarray.Length; j++)
                                        {
                                            DateTime date = Convert.ToDateTime(nstartdate);
                                            int daysToAdd = 0;// incrementDays; // 0;
                                            if (PickedDatename == "DueDate")
                                            {
                                                if (CompeltionDaysarray[j].Trim() != "")
                                                    if (flagstart == 0)
                                                    {
                                                        daysToAdd = 0;
                                                        flagstart++;
                                                        }
                                                    else
                                                    {
                                                        daysToAdd = Convert.ToInt32(CompeltionDaysarray[j]);
                                                    }
                                                if (Short_Stagelstarray[j] != "PrtFil")
                                                {
                                                    if (daysToAdd < 0)
                                                    {
                                                        while (daysToAdd < 0)
                                                        {
                                                            date = date.AddDays(-1);
                                                            DataTable dtholiday = new DataTable();
                                                            if (Session["CustomerName"].ToString()=="CUP")
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                            else
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                            if (dtholiday.Rows.Count <= 0)
                                                            {
                                                                if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                {
                                                                    daysToAdd += 1;
                                                                }
                                                            }
                                                           
                                                        }
                                                    }
                                                    else
                                                    {
                                                        while (daysToAdd > 0)
                                                        {
                                                            date = date.AddDays(1);
                                                            DataTable dtholiday = new DataTable();
                                                            if (Session["CustomerName"].ToString() == "CUP")
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                            else
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                            if (dtholiday.Rows.Count <= 0)
                                                            {
                                                                if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                {
                                                                    daysToAdd -= 1;
                                                                }
                                                            }                                                           


                                                        }
                                                    }



                                                    if (StartEndCollec.ContainsKey(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l]) == false)
                                                    {
                                                        //StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], sStartDate + "|" + date.ToString("dd-MMM-yyyy"));
                                                        StartEndCollec.Add(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], date.ToString("dd-MMM-yyyy") + "|" + date.ToString("dd-MMM-yyyy"));

                                                    }


                                                    nstartdate = date.ToString("dd-MMM-yyyy");//.ToString("MM/dd/yyyy");  // Re assign for next calculation


                                                }
                                                else
                                                {
                                                    daysToAdd = 0;
                                                    if (Short_Stagelstarray[j] == short_stage)
                                                    {
                                                        if (daysToAdd < 0)
                                                        {
                                                            while (daysToAdd < 0)
                                                            {
                                                                date = date.AddDays(-1);
                                                                DataTable dtholiday = new DataTable();
                                                                if (Session["CustomerName"].ToString() == "CUP")
                                                                    dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                                else
                                                                    dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                                if (dtholiday.Rows.Count <= 0)
                                                                {
                                                                    if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                    {
                                                                        daysToAdd += 1;
                                                                    }
                                                                }
                                                               
                                                            }
                                                        }
                                                        else
                                                        {
                                                            while (daysToAdd > 0)
                                                            {
                                                                date = date.AddDays(1);
                                                                DataTable dtholiday = new DataTable();
                                                                if (Session["CustomerName"].ToString() == "CUP")
                                                                    dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                                else
                                                                    dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                                if (dtholiday.Rows.Count <= 0)
                                                                {
                                                                    if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                    {
                                                                        daysToAdd -= 1;
                                                                    }
                                                                }
                                                               
                                                            }
                                                        }

                                                        if (StartEndCollec.ContainsKey(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l]) == false)
                                                        {
                                                            //StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], sStartDate + "|" + date.ToString("dd-MMM-yyyy"));
                                                            StartEndCollec.Add(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], date.ToString("dd-MMM-yyyy") + "|" + date.ToString("dd-MMM-yyyy"));

                                                        }

                                                        nstartdate = date.ToString("dd-MMM-yyyy");//.ToString("MM/dd/yyyy");  // Re assign for next calculation
                                                    }
                                                }
                                            }
                                            else
                                            {
                                               // if (Short_Stagelstarray[j] != "PrtFil")
                                               // {
                                                    if (Short_Stagelstarray[j] == short_stage)
                                                    {
                                                        if (daysToAdd < 0)
                                                        {
                                                            while (daysToAdd < 0)
                                                            {
                                                                date = date.AddDays(-1);
                                                            DataTable dtholiday = new DataTable();
                                                            if (Session["CustomerName"].ToString() == "CUP")
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                            else
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                            if (dtholiday.Rows.Count <= 0)
                                                            {
                                                                if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                {
                                                                    daysToAdd += 1;
                                                                }
                                                            }

                                                        }
                                                        }
                                                        else
                                                        {
                                                            while (daysToAdd > 0)
                                                            {
                                                                date = date.AddDays(1);
                                                            DataTable dtholiday = new DataTable();
                                                            if (Session["CustomerName"].ToString() == "CUP")
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where location in('US' ,'ND') and HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());
                                                            else
                                                                dtholiday = DBProc.GetResultasDataTbl("select * from tbl_Holidaylist where HolidayDate='" + date.ToString("MM-dd-yyyy") + "'", Session["sConnSiteDB"].ToString());

                                                            if (dtholiday.Rows.Count <= 0)
                                                            {
                                                                if (date.DayOfWeek != DayOfWeek.Sunday && date.DayOfWeek != DayOfWeek.Saturday)
                                                                {
                                                                    daysToAdd -= 1;
                                                                }
                                                            }
                                                            
                                                            }
                                                        }

                                                        if (StartEndCollec.ContainsKey(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l]) == false)
                                                        {
                                                            //StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], sStartDate + "|" + date.ToString("dd-MMM-yyyy"));
                                                            StartEndCollec.Add(chkchapterarray[l] + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], date.ToString("dd-MMM-yyyy") + "|" + date.ToString("dd-MMM-yyyy"));

                                                        }

                                                        nstartdate = date.ToString("dd-MMM-yyyy");//.ToString("MM/dd/yyyy");  // Re assign for next calculation
                                                    }
                                                //}
                                            }

                                        }

                                    }
                                }
                            }
                        }

                        //****** Existing logic

                        //for (int k = 0; k < Short_Stagelstarray.Length; k++)
                        //{
                        //    if (Short_Stagelstarray[k] == short_stage)
                        //    {
                        //        flag = 1;
                        //        if (flag == 1)
                        //        {

                        //            for (int j = k; j < Short_Stagelstarray.Length; j++)
                        //            {

                        //                if (dicExistsStart.ContainsKey(strChapterID + Short_Stagelstarray[j]))
                        //                {
                        //                    sStartDate = dicExistsStart[strChapterID + Short_Stagelstarray[j]];
                        //                }

                        //                //DateTime date = DateTime.ParseExact(sStartDate, "MM/dd/yyyy", null);
                        //                DateTime date = Convert.ToDateTime(sStartDate);
                        //                int daysToAdd = incrementDays; // 0;
                        //                //if (CompeltionDaysarray[j].Trim() != "")
                        //                //    daysToAdd = Convert.ToInt32(CompeltionDaysarray[j]);

                        //                if (Short_Stagelstarray[j] != "PrtFil")
                        //                {

                        //                    if (daysToAdd < 0)
                        //                    {
                        //                        while (daysToAdd < 0)
                        //                        {
                        //                            date = date.AddDays(-1);

                        //                            if (date.DayOfWeek != DayOfWeek.Sunday)  //date.DayOfWeek != DayOfWeek.Saturday
                        //                            {
                        //                                daysToAdd += 1;
                        //                            }
                        //                        }
                        //                    }
                        //                    else
                        //                    {
                        //                        while (daysToAdd > 0)
                        //                        {
                        //                            date = date.AddDays(1);

                        //                            if (date.DayOfWeek != DayOfWeek.Sunday)  //date.DayOfWeek != DayOfWeek.Saturday
                        //                            {
                        //                                daysToAdd -= 1;
                        //                            }
                        //                        }
                        //                    }

                        //                    for (int l = 0; l < chkchapterarray.Length; l++)
                        //                    {

                        //                        if (StartEndCollec.ContainsKey(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l]) == false)
                        //                        {
                        //                            //StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], sStartDate + "|" + date.ToString("dd-MMM-yyyy"));
                        //                            StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], date.ToString("dd-MMM-yyyy") + "|" + date.ToString("dd-MMM-yyyy"));

                        //                        }
                        //                    }

                        //                    // sStartDate = date.ToString("dd-MMM-yyyy");//.ToString("MM/dd/yyyy");  // Re assign for next calculation


                        //                }

                        //            }
                        //        }
                        //    }
                        //}
                        ///******






                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataSch = JsonConvert.SerializeObject(StartEndCollec) }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }

            ///** Old logic start
            ////try
            ////{

            ////    Dictionary<string, string> StartEndCollec = new Dictionary<string, string>();

            ////    List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedStageID);
            ////    List<string> chkchapterarrayIds = JsonConvert.DeserializeObject<List<string>>(chkChapter);
            ////    string[] chkchapterarray = chkchapterarrayIds.ToArray();
            ////    if (chkIds.Count > 0)
            ////    {
            ////        for (int i = 0; i < chkIds.Count; i++)
            ////        {
            ////            string strChapterID = chkIds[i].Split('|')[0];
            ////            sCompDays = chkIds[i].Split('|')[1];
            ////            string short_stage = chkIds[i].Split('|')[2];
            ////            string PickedDatename = chkIds[i].Split('|')[3];
            ////            string Short_Stagelst = chkIds[i].Split('|')[4];
            ////            string CompeltionDays = chkIds[i].Split('|')[5];
            ////            string[] Short_Stagelstarray = Short_Stagelst.Split(',');
            ////            string[] CompeltionDaysarray = CompeltionDays.Split(',');
            ////            int flag = 0;
            ////            for (int k = 0; k < Short_Stagelstarray.Length; k++)
            ////            {
            ////                if (Short_Stagelstarray[k] == short_stage)
            ////                {
            ////                    flag = 1;
            ////                    if (flag == 1)
            ////                    {
            ////                        for (int j = k; j < Short_Stagelstarray.Length; j++)
            ////                        {
            ////                            //DateTime date = DateTime.ParseExact(sStartDate, "MM/dd/yyyy", null);
            ////                            DateTime date = Convert.ToDateTime(sStartDate);
            ////                            int daysToAdd = 0;
            ////                            if (CompeltionDaysarray[j].Trim() != "")
            ////                                daysToAdd = Convert.ToInt32(CompeltionDaysarray[j]);

            ////                            while (daysToAdd > 0)
            ////                            {
            ////                                date = date.AddDays(1);

            ////                                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
            ////                                {
            ////                                    daysToAdd -= 1;
            ////                                }
            ////                            }
            ////                            for (int l = 0; l < chkchapterarray.Length; l++)
            ////                            {

            ////                                if (StartEndCollec.ContainsKey(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l]) == false)
            ////                                {
            ////                                    StartEndCollec.Add(strChapterID + "|" + Short_Stagelstarray[j] + "|" + PickedDatename + "|" + j + "|" + chkchapterarray[l], sStartDate + "|" + date.ToString("dd-MMM-yyyy"));
            ////                                }
            ////                            }
            ////                            sStartDate = date.ToString("dd-MMM-yyyy");//.ToString("MM/dd/yyyy");  // Re assign for next calculation
            ////                        }
            ////                    }
            ////                }
            ////            }


            ////        }

            ////    }
            ////    else
            ////    {
            ////        return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            ////    }

            ////    return Json(new { dataSch = JsonConvert.SerializeObject(StartEndCollec) }, JsonRequestBehavior.AllowGet);
            ////}
            ////catch (Exception ex)
            ////{
            ////    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            ////}
            //*****Old Logic end

        }
         [SessionExpire]
        public ActionResult Notes()
        {

            List<NotesModel> lst = new List<NotesModel>();

            try
            {
                DataSet ds = new DataSet();
                //  Session["sJBMAutoID"].ToString() = "BK563";
                ds = DBProc.GetResultasDataSet("Select spl.AutoArtID,spl.Instruction,spl.InstDate,(Select EmpName from JBM_EmployeeMaster where EmpAutoId=spl.EmpAutoId  or EmpLogin=spl.EmpAutoId) as[EmpName] from " + Session["sCustAcc"].ToString() + "_SplInstructions spl where spl.AutoArtID='" + Session["sJBMAutoID"].ToString() + "' ORDER BY CONVERT(DateTime, spl.InstDate,101)  DESC", Session["sConnSiteDB"].ToString());
              
                for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                {
                    DateTime instDate = DateTime.Parse(ds.Tables[0].Rows[intCount]["InstDate"].ToString());
                    string strNoteId = instDate.ToString("yyyy_MM_dd-HH-mm-ss-tt");

                    string strEmpName = "";
                    if (ds.Tables[0].Rows[intCount]["EmpName"].ToString() == "") { strEmpName = "Unknown"; } else { strEmpName = ds.Tables[0].Rows[intCount]["EmpName"].ToString(); }

                    lst.Add (new NotesModel{
                                NoteID = strNoteId + "-" + ds.Tables[0].Rows[intCount]["AutoArtID"].ToString(),
                                Date = instDate,//.ToString("MMMM dd, yyyy hh:mm:ss tt"),
                                Notes =ds.Tables[0].Rows[intCount]["Instruction"].ToString(),
                                Title ="Notes",
                                EmpAutoId = strEmpName
                    });
                }


                return View(lst);
            }
            catch (Exception)
            {
                return View(lst);
            }
        }
        [SessionExpire]
        public ActionResult InserNewNotes(string sNewNote)
        {
            try
            {
                sNewNote = HttpUtility.UrlDecode(sNewNote);
                sNewNote = sNewNote.Replace("'", "''");

                string strEmpID = "";
                if (Session["EmpIdLogin"].ToString().Length > 8) { strEmpID = "E000002"; } else { strEmpID = Session["EmpIdLogin"].ToString(); }

                string strResult = DBProc.GetResultasString("INSERT INTO " + Session["sCustAcc"].ToString() + "_SplInstructions (AutoArtID,Instruction,InstDate,DeptCode,EmpAutoId) VALUES ('" + Session["sJBMAutoID"].ToString() + "','" + sNewNote + "','" + DateTime.Now + "','|260|','" + Session["EmpAutoId"].ToString() + "')", Session["sConnSiteDB"].ToString());
                return Json(new { dataNote = "Success"}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataNote = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult UpdateDeleteNotes(string sEditedTxt, string sUniqueID, string sIsDelete)
        {
            try
            {
                string strEmpID = "";
                if (Session["EmpIdLogin"].ToString().Length > 8) { strEmpID = "E000002"; } else { strEmpID = Session["EmpIdLogin"].ToString(); }

                sEditedTxt = HttpUtility.UrlDecode(sEditedTxt);
                sEditedTxt = sEditedTxt.Replace("'", "''");
                string[] split = sUniqueID.Split('-');

                string strDate = split[0].Replace("_", "-") +  " " + split[1] + ":" + split[2] + ":" + split[3]+ ".000";
                string strAutoArtID = split[5];

                if (sIsDelete == "Delete")
                {
                    string strResult = DBProc.GetResultasString("DELETE FROM " + Session["sCustAcc"].ToString() + "_SplInstructions WHERE AutoArtID='" + strAutoArtID.Trim() + "' and InstDate = '" + strDate.Trim() + "'", Session["sConnSiteDB"].ToString());

                }
                else if (sIsDelete == "Update")
                {
                    string strResult = DBProc.GetResultasString("UPDATE " + Session["sCustAcc"].ToString() + "_SplInstructions SET Instruction = '" + sEditedTxt + "', EmpAutoId = '" + strEmpID + "' WHERE AutoArtID='" + strAutoArtID.Trim() + "' and InstDate = '" + strDate.Trim() + "'", Session["sConnSiteDB"].ToString());

                }
                return Json(new { dataNote = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataNote = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult Invoicing()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult BookIn()
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");
                string rootpath = @"" + rootdirectory + DeptFolders + "CustInput";
                //string aa= "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\"+aa;
                string folderPath = rootpath;// Server.MapPath("~/UploadedFiles/ProductionNotes");
                string XmlfolderPath = rootpath + "\\XML";//Server.MapPath("~/UploadedFiles/ProductionNotes/XML");

                gen.WriteLog("Path:" + folderPath);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                if (!Directory.Exists(XmlfolderPath))
                {
                    Directory.CreateDirectory(XmlfolderPath);
                }

                //To set the modify permission
                //clsCollec.SetFolderPermission(folderPath);
                //To set the modify permission
               // clsCollec.SetFolderPermission(XmlfolderPath);

                //Xml creation

                //<FileName>" + file.FileName + "</FileName><DisplayName>" + xmlname + "</DisplayName>
                // Save the document to a file and auto-indent the output.
                string xmlfilepath = XmlfolderPath + "\\XMLFILE.xml";
                if (!System.IO.File.Exists(xmlfilepath))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml("<Files></Files>");
                    string SavexmlPath = xmlfilepath;// Path.Combine(XmlfolderPath, "\\XMLFILE.xml");
                    doc.Save(@SavexmlPath);
                }


                BookInModel model = new BookInModel();
                List<BookInModel> dttble = new List<BookInModel>();
                string strQuery = "Select JI.JBM_ID,JI.NoofChapters,JI.JBM_Location,JI.Title,JI.BM_Author,JI.BM_ISBNnumber13,JI.BM_ISBN10number,PM.Ebook_ISBN as [eISBN],JI.JBM_Platform,JI.JBM_PrinterDate,JI.CopyrightOwner,JI.JBM_Trimsize,JI.MssPages,JI.BM_CastOffPgCount,JI.Docketno,JI.KGLAccMgrName,JI.IndexerName,JI.BM_FullService,PM.ProjectCoordInd,PM.Edition,PM.PONumber,PM.BusinessUnit,PM.ProjectManagerUS, PM.DesignLead, PM.Copyeditor, PM.ProductionLead, PM.Proofreader,PM.Current_Status,PM.Current_Health,PM.Instock,PM.SignatureSize,PM.Design,PM.FilesPath,PM.Remarks,PM.Components_Awaited,PM.Last_Review,PM.ActualPages,JI.BM_DesiredPgCount as TargetCount,PM.Launch,PM.misc1,PM.misc2,PM.misc3,PM.misc4 from JBM_Info JI JOIN " + Session["sCustAcc"].ToString() + "_ProjectManagement PM ON JI.JBM_AutoID = PM.JBM_AutoID where JI.JBM_AutoID='" + Session["sJBMAutoID"].ToString() + "'";
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    model.StageService = "";// Convert.ToString(ds.Tables[0].Rows[0]["JBM_ID"].ToString().Trim());                                    


                    //Load PagingApp list Items
                    List<SelectListItem> items = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select PlatformID,PlatformDesc from JBM_Platform", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            items.Add(new SelectListItem
                            {
                                Text = myRow["PlatformDesc"].ToString(),
                                Value = myRow["PlatformID"].ToString()
                            });
                        }
                    }
                    ViewBag.Platformlist = items;

                    //Load Stage list Items
                    List<SelectListItem> stageitems = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select Stages from [dbo].[JBM_AccountTypeDesc] WHERE CustAccess='BK'", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string strTemp = ds.Tables[0].Rows[0]["Stages"].ToString();
                        string[] strIn = strTemp.Split('|'); //Split(strTemp, "|", -1, CompareMethod.Binary)

                        for (int i = 0; i < strIn.Length; i++)
                        {
                            string[] strInvalues = strIn[i].Split('-');
                            if (strInvalues.Length > 1)
                            {
                                string colval = strInvalues[0].ToString();
                                string dataval = strInvalues[1].ToString();
                                stageitems.Add(new SelectListItem
                                {
                                    Text = colval.ToString(),
                                    Value = dataval.ToString()
                                });
                            }
                        }

                    }
                    ViewBag.Stagelist = stageitems;

                    //Load Complexity list Items
                    List<SelectListItem> Complexityitems = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("SELECT ComplexityId,ComplexityDesc from JBM_Complexity", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            Complexityitems.Add(new SelectListItem
                            {
                                Text = myRow["ComplexityDesc"].ToString(),
                                Value = myRow["ComplexityId"].ToString()
                            });
                        }
                    }
                    ViewBag.Complexitylist = Complexityitems;

                    //Load ChpType list Items
                    List<SelectListItem> ChpTypeitems = new List<SelectListItem>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select ArtTypeID,ArtTypeDesc from JBM_ArticleTypes", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            ChpTypeitems.Add(new SelectListItem
                            {
                                Text = myRow["ArtTypeDesc"].ToString(),
                                Value = myRow["ArtTypeID"].ToString()
                            });
                        }
                    }
                    ViewBag.ChpTypelist = ChpTypeitems;

                    DataTable dt = BindGrid(true);
                    ViewBag.dt = dt;
                }
                return View(model);
            }
            catch (Exception ex){
                gen.WriteLog("Error: " + ex.Message);
                ViewBag.Error = ex.Message;
                return View();
            }
        }
        [SessionExpire]
        public ActionResult Book_In()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult ProductionNotes()
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }

                rootdirectory= rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");

                string rootpath = @"" + rootdirectory + DeptFolders + "ProductionNotes";
               


                //string aa= "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\"+aa;
                string folderPath = rootpath;// Server.MapPath("~/UploadedFiles/ProductionNotes");
                string XmlfolderPath = rootpath + "\\XML";//Server.MapPath("~/UploadedFiles/ProductionNotes/XML");


                gen.WriteLog("Path:" + folderPath);
                try
                {
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    if (!Directory.Exists(XmlfolderPath))
                    {
                        Directory.CreateDirectory(XmlfolderPath);
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("Directory Error:" + ex.Message);
                }

                //To set the modify permission
                //clsCollec.SetFolderPermission(folderPath);

                //To set the modify permission
                //clsCollec.SetFolderPermission(XmlfolderPath);
                //Xml creation

                //<FileName>" + file.FileName + "</FileName><DisplayName>" + xmlname + "</DisplayName>
                // Save the document to a file and auto-indent the output.
                string xmlfilepath = XmlfolderPath + "\\XMLFILE.xml";
                List<FileModel> files = new List<FileModel>();
                try
                {
                    if (!System.IO.File.Exists(xmlfilepath))
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml("<Files></Files>");
                        string SavexmlPath = xmlfilepath;// Path.Combine(XmlfolderPath, "\\XMLFILE.xml");
                        doc.Save(@SavexmlPath);
                    }

                    //Fetch all files in the Folder (Directory).
                    //string[] filePaths = Directory.GetFiles(Server.MapPath("~/UploadedFiles/ProductionNotes/XML"));

                    //Copy File names to Model collection.
                   
                    string xmlfile = xmlfilepath; // Server.MapPath("~/UploadedFiles/ProductionNotes/XML/XMLFILE.xml");
                    if (xmlfile != "")
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(xmlfile);
                        XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Files/File");

                        foreach (XmlNode node in nodeList)
                        {
                            string filedate = node.SelectSingleNode("Date").InnerText.ToString();
                            //filedate.ToString("MM/dd/yyyy");
                            string exenam = node.SelectSingleNode("FileName").InnerText.Split('.')[1] + ".png";
                            if (System.IO.File.Exists(folderPath + @"\" + node.SelectSingleNode("FileName").InnerText))
                            {
                                files.Add(new FileModel { FileName = node.SelectSingleNode("FileName").InnerText, DisplayName = node.SelectSingleNode("DisplayName").InnerText, UploadDate = filedate, exename = exenam });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("XML Load Error:" + ex.Message);
                }

                return View(files);
            }
            catch (Exception ex)
            {
                gen.WriteLog("Error: " + ex.Message);
                return View();
            }
        }

        [HttpPost]
        [SessionExpire]
        public ActionResult ProductionNotesUpload(List<HttpPostedFileBase> fileUpload, string[] FileData)
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }

                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");

                string rootpath = @"" + rootdirectory + DeptFolders + "ProductionNotes";
                //To set the modify permission
                //clsCollec.SetFolderPermission(rootpath);

                //string aa = "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\" + aa;
                List<FileModel> files = new List<FileModel>();
                if (fileUpload != null)
                {
                    int i = 0;
                    foreach (HttpPostedFileBase file in fileUpload)
                    {

                        if (file != null)
                        {
                            string SavefilePath = rootpath + "\\" + file.FileName;
                            if (!System.IO.File.Exists(SavefilePath))
                            {
                               
                                System.IO.File.WriteAllBytes(SavefilePath, clsCollec.ReadData(file.InputStream));

                                //add xml
                                XmlDocument xmlFileDoc = new XmlDocument();
                                xmlFileDoc.Load(rootpath + "\\XML\\XMLFILE.xml");
                                XmlElement ParentElement = xmlFileDoc.CreateElement("File");
                                XmlElement FileName = xmlFileDoc.CreateElement("FileName");
                                FileName.InnerText = file.FileName;
                                XmlElement DisplayName = xmlFileDoc.CreateElement("DisplayName");
                                string strDesc = "";
                                if (FileData[i] == "")
                                { strDesc = ""; }
                                else {
                                    strDesc = FileData[i].ToString();
                                }
                                DisplayName.InnerText = strDesc;
                                XmlElement Date = xmlFileDoc.CreateElement("Date");
                                Date.InnerText = DateTime.Now.ToString("dd-MMM-yyyy");
                                ParentElement.AppendChild(FileName);
                                ParentElement.AppendChild(DisplayName);
                                ParentElement.AppendChild(Date);

                                xmlFileDoc.DocumentElement.AppendChild(ParentElement);
                                xmlFileDoc.Save(rootpath + "\\XML\\XMLFILE.xml");
                                string filedate = DateTime.Now.ToString("dd-MMM-yyyy"); // System.IO.File.GetLastWriteTime(SavefilePath).ToShortDateString();
                                //filedate.ToString("MM/dd/yyyy");
                                files.Add(new FileModel { FileName = file.FileName, DisplayName = strDesc, UploadDate = filedate });
                            }
                            else
                            {
                                files.Add(new FileModel { FileName = file.FileName, DisplayName = "Already exists" });
                                //return Json(files, JsonRequestBehavior.AllowGet);
                            }
                        }
                        i++;
                    }


                }
                return Json(files, JsonRequestBehavior.AllowGet);
                //return RedirectToAction("ProductionNotes");

            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }


        }
        public FileResult DownloadFile(string fileName)
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");
                string rootpath = @"" + rootdirectory + DeptFolders + "ProductionNotes";
                //To set the modify permission
                //clsCollec.SetFolderPermission(rootpath);
                //string aa = "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\" + aa;
                byte[] bytes;
                string path = "";

                path = rootpath + "\\" + fileName;
                bytes = System.IO.File.ReadAllBytes(path);
                return File(bytes, "application/octet-stream", fileName);
            }
            catch (Exception)
            {

                throw;
            }
        }
        [SessionExpire]
        public ActionResult ShowFile(string fileName)
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");
                string rootpath = @"" + rootdirectory + DeptFolders + "ProductionNotes";
                //To set the modify permission
               // clsCollec.SetFolderPermission(rootpath);
                //string aa = "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\" + aa;
                string assemblyPath = rootpath; //System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);

                Process proc = new Process();
                proc.StartInfo = new ProcessStartInfo()
                {
                    FileName = assemblyPath + "\\" + fileName
                };
                proc.Start();
                return RedirectToAction("ProductionNotes");
            }
            catch (Exception)
            {

                throw;
            }
        }
        [SessionExpire]
        public ActionResult DeleteFile(string fileName)
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");
                string rootpath = @"" + rootdirectory + DeptFolders + "ProductionNotes";

                //To set the modify permission
                clsCollec.SetFolderPermission(rootpath);

                //'' The below code is working and need to add web config  <identity impersonate="true"/>
                //DirectorySecurity sec = Directory.GetAccessControl(rootpath);
                //SecurityIdentifier everyone = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
                //sec.AddAccessRule(new FileSystemAccessRule(everyone, FileSystemRights.Modify | FileSystemRights.Synchronize, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
                //Directory.SetAccessControl(rootpath, sec);

                // string rootpath = @"\\nodnas03\cenpro\ApplicationFiles\books\TandF\CHARLES2_9781138634732-200023\Input\CustInput\ProductionNotes";
                string flname = "";
                // string[] filePaths = Directory.GetFiles(Server.MapPath("~/UploadedFiles/ProductionNotes/XML/"));
                // byte[] bytes;
                string path = "";
                //Copy File names to Model collection.

                string xmlfile = rootpath + "\\XML\\XMLFILE.xml";
                if (xmlfile != "")
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(xmlfile);
                    XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Files/File");

                    foreach (XmlNode node in nodeList)
                    {
                        if (node.SelectSingleNode("FileName").InnerText == Convert.ToString(fileName))
                        {
                            flname = node.SelectSingleNode("FileName").InnerText;
                            path = rootpath + "\\" + flname;
                            System.IO.File.SetAttributes(path, FileAttributes.Normal);
                            System.IO.File.Delete(path);
                            node.ParentNode.RemoveChild(node);
                            xmlDoc.Save(xmlfile);
                        }
                    }
                }


                return RedirectToAction("ProductionNotes");
            }
            catch (Exception)
            {

                throw;
            }
        }
        [HttpPost]
        [SessionExpire]
        public ActionResult BookInUpload(List<HttpPostedFileBase> fileUpload, string[] FileData)
        {
            try
            {
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                rootdirectory = rootdirectory.Replace(@"\\nodnas03\", @"\\nodnas03.kwglobal.com\");

                string rootpath = @"" + rootdirectory + DeptFolders + "CustInput";

               //To set permission 
                clsCollec.SetFolderPermission(rootpath);

                //string aa = "\\10.18.2.48\\d$\\SmartTrack_Changes\\new\\ProductionNotes";
                //string rootpath = @"\" + aa;
                List<FileModel> files = new List<FileModel>();
                if (fileUpload != null)
                {
                    int i = 0;
                    foreach (HttpPostedFileBase file in fileUpload)
                    {

                        if (file != null)
                        {
                            string SavefilePath = rootpath + "\\" + file.FileName;
                            if (!System.IO.File.Exists(SavefilePath))
                            {

                                System.IO.File.WriteAllBytes(SavefilePath, clsCollec.ReadData(file.InputStream));

                                //add xml
                                XmlDocument xmlFileDoc = new XmlDocument();
                                xmlFileDoc.Load(rootpath + "\\XML\\XMLFILE.xml");
                                XmlElement ParentElement = xmlFileDoc.CreateElement("File");
                                XmlElement FileName = xmlFileDoc.CreateElement("FileName");
                                FileName.InnerText = file.FileName;
                                XmlElement DisplayName = xmlFileDoc.CreateElement("DisplayName");
                                string strDesc = "";
                                if (FileData[i] == "undefined") { strDesc = ""; } else { strDesc = FileData[i]; }
                                DisplayName.InnerText = strDesc;
                                XmlElement Date = xmlFileDoc.CreateElement("Date");
                                Date.InnerText = DateTime.Now.ToString("dd-MMM-yyyy");
                                ParentElement.AppendChild(FileName);
                                ParentElement.AppendChild(DisplayName);
                                ParentElement.AppendChild(Date);

                                xmlFileDoc.DocumentElement.AppendChild(ParentElement);
                                xmlFileDoc.Save(rootpath + "\\XML\\XMLFILE.xml");
                                string filedate = DateTime.Now.ToString("dd-MMM-yyyy"); // System.IO.File.GetLastWriteTime(SavefilePath).ToShortDateString();
                                //filedate.ToString("MM/dd/yyyy");
                                files.Add(new FileModel { FileName = file.FileName, DisplayName = strDesc, UploadDate = filedate });
                            }
                            else
                            {
                                files.Add(new FileModel { FileName = file.FileName, DisplayName = "Already exists" });
                                //return Json(files, JsonRequestBehavior.AllowGet);
                            }
                        }
                        i++;
                    }


                }
                return Json(files, JsonRequestBehavior.AllowGet);


            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }


        }

        public string Proc_IntrnlID(string strJAutoId)
        {
            SqlConnection sqlConn = new SqlConnection();
            string Result = string.Empty;
            // Proc_IntrnlID = "";
            DataTable clsFetchArticleInfo = new DataTable();
            try
            {
                DateTime now = DateTime.Today;
                string year = now.ToString("yyyy");
                string yearmid = year.Substring(2, 2);
                string qrystr = "Select  MAX(IntrnlID) as IntrnlID  from " + Session["sCustAcc"].ToString() + "_chapterinfo WHERE JBM_AutoId='" + strJAutoId + "' AND IntrnlID like '" + yearmid + "%'";
                clsFetchArticleInfo = DBProc.GetResultasDataTbl(qrystr, Session["sConnSiteDB"].ToString());
                string intTempID = "";

                if (clsFetchArticleInfo.Rows.Count > 0)
                {
                    if (clsFetchArticleInfo.Rows[0]["IntrnlID"].ToString() != "")
                    {
                        intTempID = clsFetchArticleInfo.Rows[0]["IntrnlID"].ToString();
                        intTempID = String.Format("{0, 0:D6}", Convert.ToInt64(intTempID) + 1); //String.Format("Number {0, 0:D5}", num)
                    }
                    else
                    {
                        intTempID = DateTime.Now.Year.ToString().Substring(2, 2) + "0001";
                    }
                }
                Result = intTempID;
            }
            catch (Exception ex)
            {
                Result = "Error";
            }
            finally
            {
                sqlConn.Close();
            }


            return Result;

        }
        public DataTable SetRecord_All(string strFieldVal)
        {
            DataTable strOutRecordCol = new DataTable();
            strOutRecordCol = Proc_Set_FieldVal(strFieldVal);
            return strOutRecordCol;
        }
        public DataTable Proc_Set_FieldVal(string strTemp)
        {
            DataTable strOut = new DataTable();
            //Dim FieldVal As Init_ColVal
            //FieldVal = Nothing

            string[] strIn = strTemp.Split('|'); //Split(strTemp, "|", -1, CompareMethod.Binary)

            try
            {
                DataRow dtrow = strOut.NewRow();// DataRow();
                strOut.Rows.Add(dtrow);
                for (int i = 0; i < strIn.Length; i++)
                {
                    //    FieldVal = Nothing  
                    string[] strInvalues = strIn[i].Split(',');
                    //    FieldVal.gColName = Mid(strIn(i), 1, InStr(1, strIn(i), ",", CompareMethod.Binary) - 1) 
                    //    Dim strVal As String = Mid(strIn(i), InStr(1, strIn(i), ",", CompareMethod.Binary) + 1)   
                    if (strInvalues.Length > 0)
                    {
                        string colval = strInvalues[0].ToString();
                        string dataval = strInvalues[1].ToString();
                        strOut.Columns.Add(colval, typeof(string));
                        DataRow dr = strOut.Rows[0];
                        dr[colval] = dataval;

                    }

                }
            }
            catch (Exception ex)
            {
                // Result = "Error";
            }
            finally
            {
                //FieldVal = Nothing;
            }

            return strOut;
        }
        public static string dt_DateFrmt(DateTime dtDateInput)
        {
            string strTemp = null;
            if (dtDateInput != null)
                strTemp = String.Format("{0:D2}", dtDateInput.Month) + "-" + String.Format("{0:D2}", dtDateInput.Day) + "-" + dtDateInput.Year;
            return strTemp;
        }
        [HttpPost]

        public JsonResult UpdateBookIn(BookInModel obj)
        {
            try
            {
                if (obj.ChapterRange == null)
                {
                    return Json(new { dataSch = "Range" }, JsonRequestBehavior.AllowGet);
                }
                else { insertData(obj); }
               
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public void insertData(BookInModel obj)
        {
            SqlConnection con = new SqlConnection();
            con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            con.Open();
            string sCustJobIDquery = "Select Job_ID from JBM_AccountTypeDesc where CustAccess = 'BK'";
            DataTable dtcustid = new DataTable();
            dtcustid = DBProc.GetResultasDataTbl(sCustJobIDquery, Session["sConnSiteDB"].ToString());

            if (obj.Batch != "" && obj.Batch != null)
            {
                obj.Batch = obj.Batch.Replace("Batch", "");
            }
            string strAutoIdFilter = dtcustid.Rows[0]["Job_ID"].ToString();
            string strStage = "";
            strStage = obj.StageService.ToString();
            DataTable clsCustSave = new DataTable();
            string strContcheck = "";
          
            string[] strChapterNames = obj.ChapterRange.Split(',');
            bool blnDuplicateFound = false;
            string strJAutoID = Session["sJBMAutoID"].ToString();
            string strUserID = Session["EmpIdLogin"].ToString();
            if (obj.StageService.ToString().ToUpper().Trim() == "SAMPLE")
            {
                if (strChapterNames.Length > 0)
                {
                    Regex Regex1 = new Regex(@"([0-9]\-[0-9])", RegexOptions.IgnoreCase);
                    Match m = Regex1.Match(strChapterNames[0].ToString());
                    if (m.Success)
                    {
                        //lbl_Error.Text = "More than one chapter<br>"
                    }
                    else
                    {
                        //Regex1 =Nothing;
                        Regex1 = new Regex(@"([a-zA-Z])", RegexOptions.IgnoreCase);
                        Match m1 = Regex1.Match(strChapterNames[0].ToString());
                        if (m1.Success)
                        {
                            strContcheck += strChapterNames[0] + char.ConvertFromUtf32((int)13);
                        }
                        else
                        {
                            strContcheck += "Chp" + String.Format("{0:D2}", Convert.ToInt32(strChapterNames[0])) + char.ConvertFromUtf32((int)13);
                        }
                    }
                }
                else
                {
                    //lbl_Error.Text = "More than one chapter<br>"               
                }
            }
            else
            {                //lbl_Error.Text = ""
                for (int i = 0; i < strChapterNames.Length ; i++)
                {
                    if (strChapterNames.Length > 0)
                    {
                        Regex Regex2 = new Regex(@"([0-9]\-[0-9])", RegexOptions.IgnoreCase);
                        Match m2 = Regex2.Match(strChapterNames[i].ToString());
                        bool blnAlphaCheck1 = false;

                        if (m2.Success)
                        {
                            string[] strChapnumbers = strChapterNames[i].ToString().Split('-');
                            int result;
                            if (int.TryParse(strChapnumbers[0], out result) && int.TryParse(strChapnumbers[1], out result))
                            {
                                int chpStart = Convert.ToInt32(strChapnumbers[0]);
                                int chpEnd = Convert.ToInt32(strChapnumbers[1]);
                                for (int j = chpStart; j <= chpEnd; j++)
                                {
                                    //string searchstring = "Chp" + String.Format("{0:D2}");

                                    if (strContcheck != "")
                                    {
                                        if (strContcheck.IndexOf("Chp" + String.Format("{0:D2}", j), 1) > 0)
                                        {
                                            blnDuplicateFound = true;
                                            //return;
                                        }
                                    }
                                    strContcheck += "Chp" + String.Format("{0:D2}", j) + char.ConvertFromUtf32((int)13);
                                }
                            }
                            else
                            {
                                blnAlphaCheck1 = true;
                            }
                        }
                        else
                        {
                            blnAlphaCheck1 = true;
                        }


                        if (blnAlphaCheck1 == true)
                        {
                            // Regex1 = Nothing

                            Regex Regex3 = new Regex(@"([a-zA-Z])", RegexOptions.IgnoreCase);
                            Match m3 = Regex3.Match(strChapterNames[i].ToString());
                            if (m3.Success)
                            {
                                if (strContcheck != "")
                                {
                                    if (strContcheck.IndexOf(strChapterNames[i], 1) > 0)
                                    {
                                        blnDuplicateFound = true;
                                        break;
                                        //Exit For
                                    }
                                }
                               
                                strContcheck += strChapterNames[i] + char.ConvertFromUtf32((int)13);


                               
                            }
                            else
                            {
                                if (strContcheck != "")
                                {
                                    if (strContcheck.IndexOf("Chp" + String.Format("{0:D2}", Convert.ToInt32(strChapterNames[i])), 1) > 0) //InStr(1, strContcheck, "Chp" + String.Format("{0:D2}", Convert.ToInt32(strChapterNames[i])), Microsoft.VisualBasic.CompareMethod.Text) > 0)
                                    {
                                        blnDuplicateFound = true;
                                        break;
                                        // Exit For
                                    }
                                    strContcheck += "Chp" + String.Format("{0:D2}", Convert.ToInt32(strChapterNames[i])) + char.ConvertFromUtf32((int)13);
                                }
                                else
                                {
                                    strContcheck += "Chp" + String.Format("{0:D2}", Convert.ToInt32(strChapterNames[i])) + char.ConvertFromUtf32((int)13);
                                }
                                
                            }
                        }
                    }



                }

            }
            string strProofType = "";
            string WFlwQuery;
            string WFlow = "";
            string WFcodesplit = "";
            string WFStart = "";
            string WFEnd = "";
            bool strJBM_Flow = false;
            string ops_wf = "";
            string StrCorrAut = "";
            string strIntrnlJid = "";

            WFlwQuery = "Select " + strStage + "_wF, JBM_ID, JBM_Flow, JBM_TeamID, JBM_ProofType,OPS_wf,BM_Author from JBM_Info where JBM_AutoID = '" + strJAutoID + "' and JBM_Disabled='0'";
            DataTable Dttemp1 = new DataTable();
            Dttemp1 = DBProc.GetResultasDataTbl(WFlwQuery, Session["sConnSiteDB"].ToString());

            if (Dttemp1.Rows[0]["JBM_Flow"].ToString() != null && Dttemp1.Rows[0]["JBM_Flow"].ToString() != "")
                strJBM_Flow = false;
            else
                strJBM_Flow = true;

            if (Dttemp1.Rows[0][strStage + "_wF"].ToString() != null && Dttemp1.Rows[0][strStage + "_wF"].ToString() != "")
                WFlow = Dttemp1.Rows[0][strStage + "_wF"].ToString();
            else
                WFlow = "W1";

            if (Dttemp1.Rows[0]["JBM_ProofType"].ToString() != null && Dttemp1.Rows[0]["JBM_ProofType"].ToString() != "")
                strProofType = Dttemp1.Rows[0]["JBM_ProofType"].ToString();

            if (Dttemp1.Rows[0]["OPS_wf"].ToString() != null && Dttemp1.Rows[0]["OPS_wf"].ToString() != "")
                ops_wf = Dttemp1.Rows[0]["OPS_wf"].ToString();

            if (Dttemp1.Rows[0]["BM_Author"].ToString() != null && Dttemp1.Rows[0]["BM_Author"].ToString() != "")
                StrCorrAut = Dttemp1.Rows[0]["BM_Author"].ToString();

            if (Dttemp1.Rows[0]["JBM_ID"].ToString() != null && Dttemp1.Rows[0]["JBM_ID"].ToString() != "")
                strIntrnlJid = Dttemp1.Rows[0]["JBM_ID"].ToString();
            else
                strIntrnlJid = "";

            WFlwQuery = "Select WFCode, CeParallel, XMLParallel from JBM_WFCode where WFNAme = '" + WFlow + "'";
            Dttemp1 = DBProc.GetResultasDataTbl(WFlwQuery, Session["sConnSiteDB"].ToString());

            if (Dttemp1.Rows[0]["WFCode"].ToString() != null && Dttemp1.Rows[0]["WFCode"].ToString() != "")
            {
                string[] WFcodestage = Dttemp1.Rows[0]["WFCode"].ToString().Split('|');
                WFcodesplit = WFcodestage[0];
                string[] WFCodeSplitStage = WFcodesplit.Split(',');
                WFStart = WFCodeSplitStage[0];
                WFEnd = WFCodeSplitStage[1];
                if (WFEnd.Length == 1)
                    WFEnd = 0 + WFEnd;
            }

            string strCeParallel = "";
            string strXMLParallel = "";

            if (Dttemp1.Rows[0]["CeParallel"].ToString() != null && Dttemp1.Rows[0]["CeParallel"].ToString() != "")
                strCeParallel = Dttemp1.Rows[0]["CeParallel"].ToString();
            else
                strCeParallel = "";


            if (Dttemp1.Rows[0]["XMLParallel"].ToString() != null && Dttemp1.Rows[0]["XMLParallel"].ToString() != "")
                strXMLParallel = Dttemp1.Rows[0]["XMLParallel"].ToString();
            else
                strXMLParallel = "";

            string[] ChapterNames = strContcheck.Split(Convert.ToChar(char.ConvertFromUtf32((int)13)));
            string strFieldVal = "";
            int intAutoId;

            for (int i = 0; i < ChapterNames.Length - 1; i++)
            {
                string strChapterID = "";

                if (ChapterNames[i] != "")
                {
                    if (obj.HasVolumeDetails == true)
                        strChapterID = "Vol-" + obj.Volume + "-" + ChapterNames[i].ToString().Trim();
                    else
                        strChapterID = ChapterNames[i].ToString().Trim();

                    string strAutoArtIDTemp = "";
                    string strArtStage = "";
                    string strGraphArtStage = "";

                    if (strStage == "FP")
                    {
                        strArtStage = "S100" + WFStart + "_S1009_S100" + WFStart + "_40385_";
                    }
                    else
                    {
                        if (WFEnd.Length == 3)
                            strArtStage = "S100" + WFStart + "_S1010_S1" + WFEnd + "_40385_";
                        else
                            strArtStage = "S100" + WFStart + "_S1010_S10" + WFEnd + "_40385_";
                    }

                    strGraphArtStage = "S1146_S1009_S1146_40385_";

                    clsCustSave = new DataTable();
                    clsCustSave = DBProc.GetResultasDataTbl("Select max(AutoID) as AutoID from  " + Session["sCustAcc"].ToString() + "_chapterinfo", Session["sConnSiteDB"].ToString());

                    if (clsCustSave.Rows[0]["AutoID"] != null)
                    {
                        intAutoId = Convert.ToInt32(clsCustSave.Rows[0]["AutoID"].ToString());
                        intAutoId += 1;
                        strAutoArtIDTemp = intAutoId.ToString();

                        switch (strAutoArtIDTemp.Length)
                        {
                            case 1:
                                strAutoArtIDTemp = strAutoIdFilter + "000" + strAutoArtIDTemp;
                                break;
                            case 2:
                                strAutoArtIDTemp = strAutoIdFilter + "00" + strAutoArtIDTemp;
                                break;
                            case 3:
                                strAutoArtIDTemp = strAutoIdFilter + "0" + strAutoArtIDTemp;
                                break;
                            default:
                                strAutoArtIDTemp = strAutoIdFilter + strAutoArtIDTemp;
                                break;
                        }
                    }
                    else
                    {
                        strAutoArtIDTemp = strAutoIdFilter + "0001";
                        intAutoId = 1;
                    }

                    string strIntnlID = Proc_IntrnlID(strJAutoID);


                    clsCustSave = new DataTable();
                    strFieldVal = "AutoID," + intAutoId + "|";
                    strFieldVal += "AutoArtID," + strAutoArtIDTemp.Trim() + "|";// 'Chapterinfo table and Revprodinfo
                    if (strStage != "Sample" && strStage != "FP" && strStage != "TxtA" && strStage != "ImgA" && strStage != "PhotoR")
                        strFieldVal += "FpRevPtrInfo," + strStage + "1|";
                    else
                        strFieldVal += "FpRevPtrInfo," + strStage + "|";


                    strFieldVal += "IntrnlID," + strIntnlID + "|";
                    strFieldVal += "JBM_AutoID," + strJAutoID + "|";// 'Chapter info table JBMAutoid
                    strFieldVal += "ArtTypeID,0|";//  'chaptertype  RevProductonInfo
                    strFieldVal += "PlatformID," + obj.Platform + "|";// cbo_Platform.SelectedItem.Value + "|";//  ' Chapter info tableplatform
                    strFieldVal += "ChapterID," + strChapterID.Trim() + "|";// 'Chapter info table chaptername
                    strFieldVal += "DOI," + strChapterID.Trim().Replace("'", "''") + "|";
                    strFieldVal += "DiskInfo,1|";// 'chapter info table
                    strFieldVal += "Eproof,0|";// ' Chapter info table bydefault 0 
                    strFieldVal += "ModeofInput,0|";// 'chapter info table
                    strFieldVal += "Priority,0|";// 'Chapter info table
                    strFieldVal += "Complexity," + obj.Complexity + "|";// cbo_Complexity.SelectedItem.Value + "|";//   ' chapter info table 
                    strFieldVal += "Active,1" + "|";// 'Chapter info table
                    strFieldVal += "WIP,1|";
                    strFieldVal += "LoginEnteredDate," + DateTime.Now.ToString("MM/dd/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + "|";
                    strFieldVal += "LoginRecDate," + DateTime.Now.ToString("MM/dd/yyyy") + " " + System.DateTime.Now.ToString("hh:mm:ss") + "|";
                    strFieldVal +=  obj.ReceiveedDate.ToString().Trim()!=""? "RecDate,'" + Convert.ToDateTime(obj.ReceiveedDate.ToString()).ToString("MM/dd/yyyy") + "'|" : "RecDate,"  + "null|";
                    strFieldVal += obj.DueDate.ToString().Trim() != "" ? "DueDate,'" + Convert.ToDateTime(obj.DueDate.ToString()).ToString("MM/dd/yyyy") + "'|" : "DueDate,"  + "null|";
                    strFieldVal +=  obj.CustDueDate.ToString().Trim() != "" ? "customerrecdate,'" + Convert.ToDateTime(obj.CustDueDate.ToString()).ToString("MM/dd/yyyy") + "'|" : "customerrecdate,"  + "null|";
                    strFieldVal += obj.CustDueDate.ToString().Trim() != "" ? "customerduedate,'" + Convert.ToDateTime(obj.CustDueDate.ToString()).ToString("MM/dd/yyyy") + "'|" : "customerduedate,"  + "null|";
                    strFieldVal +=  obj.RevisedDate.ToString().Trim() != "" ? "JNRevisedDate,'" + Convert.ToDateTime(obj.RevisedDate.ToString()).ToString("MM/dd/yyyy") + "'|" : "JNRevisedDate,"  + "null|";// '' rev add by maya
                    strFieldVal += "Rev_wf," + WFlow + "|";
                    strFieldVal += "Ceparallel," + strCeParallel + "|";
                    strFieldVal += "XMLparallel," + strXMLParallel + "|";
                    strFieldVal += "Loginby," + Session["EmpAutoId"].ToString().Replace("'", "''") + "|";  //strUserID
                    strFieldVal += "CWP_Article," + strProofType + "|";
                    strFieldVal += "OPS_wf," + ops_wf + "|";
                    strFieldVal += "CorrAut," + StrCorrAut.Replace("'", "''") + "|";

                    if (obj.Batch != "")
                        strFieldVal += "Batch," + obj.Batch + "|";
                    else
                        strFieldVal += "Batch,|";

                    //''S1146_S1009_S1146_40385
                    strFieldVal += "ArtStageTypeID," + strArtStage + dt_DateFrmt(Convert.ToDateTime(DateTime.Now.ToShortDateString())) + "_" + DateTime.Now.ToLongTimeString() + "|";
                    strFieldVal += "GraphArtStageTypeID," + strGraphArtStage + dt_DateFrmt(Convert.ToDateTime(DateTime.Now.ToShortDateString())) + "_" + DateTime.Now.ToLongTimeString() + "|";
                    clsCustSave = SetRecord_All(strFieldVal);

                    DataTable dt = new DataTable();
                    DataTable DtStageInfo = new DataTable();
                    dt = DBProc.GetResultasDataTbl("Select * from " + Session["sCustAcc"].ToString() + "_chapterinfo WHERE (ChapterID='" + strChapterID.Trim() + "') AND (JBM_AutoID='" + strJAutoID + "')", Session["sConnSiteDB"].ToString());
                    string strAutoArtID = "";
                    if (dt.Rows.Count == 1)
                    {
                        strAutoArtID = dt.Rows[0]["AutoArtId"].ToString();
                    }
                    else
                    {
                        strAutoArtID = "";
                    }
                    if (dt.Rows.Count == 0)
                    {
                        string strinsertchapterinfo = "Insert into  " + Session["sCustAcc"].ToString() + "_chapterinfo(AutoID,AutoArtID,JBM_AutoID,PlatformID,ChapterID,Complexity,DiskInfo,Eproof,ModeofInput,Priority,Active,WIP,Batch,ArtTypeID, IntrnlID,CorrAut,DOI) Values('" + clsCustSave.Rows[0]["AutoID"].ToString() + "','" + clsCustSave.Rows[0]["AutoArtID"].ToString() + "','" + clsCustSave.Rows[0]["JBM_AutoID"].ToString() + "','" + clsCustSave.Rows[0]["PlatformID"].ToString() + "','" + clsCustSave.Rows[0]["ChapterID"].ToString() + "','" + clsCustSave.Rows[0]["Complexity"].ToString() + "','" + clsCustSave.Rows[0]["DiskInfo"].ToString() + "','" + clsCustSave.Rows[0]["Eproof"].ToString() + "','" + clsCustSave.Rows[0]["ModeofInput"].ToString() + "','" + clsCustSave.Rows[0]["Priority"].ToString() + "','" + clsCustSave.Rows[0]["Active"].ToString() + "','" + clsCustSave.Rows[0]["WIP"].ToString() + "','" + clsCustSave.Rows[0]["Batch"].ToString() + "','" + clsCustSave.Rows[0]["ArtTypeID"].ToString() + "','" + clsCustSave.Rows[0]["IntrnlID"].ToString() + "','" + clsCustSave.Rows[0]["CorrAut"].ToString() + "','" + clsCustSave.Rows[0]["DOI"].ToString() + "')";
                        string strinsertProdInfo = "Insert into  " + Session["sCustAcc"].ToString() + "_ProdInfo(AutoArtID,FpRevPtrInfo,JNRevisedDate,CWP_Article,OPS_wf) values ('" + clsCustSave.Rows[0]["AutoArtID"].ToString() + "','" + clsCustSave.Rows[0]["FpRevPtrInfo"].ToString() + "' ," +clsCustSave.Rows[0]["JNRevisedDate"].ToString().Trim() + ",'" + clsCustSave.Rows[0]["CWP_Article"].ToString() + "','" + clsCustSave.Rows[0]["OPS_wf"].ToString() + "')";
                        string strinsertStageinfo = "Insert into " + Session["sCustAcc"].ToString() + "_Stageinfo(AutoArtId, RevFinStage, ArtStageTypeID, Rev_wf, ReceivedDate, DueDate, Loginby, LoginEnteredDate, LoginRecDate, Ceparallel, XMLParallel,CustomerRecDate,CustomerDueDate,GraphArtStageTypeID) values ('" + clsCustSave.Rows[0]["AutoArtID"].ToString() + "', '" + clsCustSave.Rows[0]["FpRevPtrInfo"].ToString() + "', '" + clsCustSave.Rows[0]["ArtStageTypeID"].ToString() + "', '" + clsCustSave.Rows[0]["Rev_wf"].ToString() + "', " +clsCustSave.Rows[0]["RecDate"].ToString().Trim() + ", " +clsCustSave.Rows[0]["DueDate"].ToString() + ", '" + clsCustSave.Rows[0]["Loginby"].ToString() + "', '" + clsCustSave.Rows[0]["LoginEnteredDate"].ToString()+ "', '" + clsCustSave.Rows[0]["LoginRecDate"].ToString() + "','" + clsCustSave.Rows[0]["CeParallel"].ToString() + "','" + clsCustSave.Rows[0]["XMLParallel"].ToString() + "'," + clsCustSave.Rows[0]["customerrecdate"].ToString() + "," + clsCustSave.Rows[0]["customerduedate"].ToString() + ", '" + clsCustSave.Rows[0]["GraphArtStageTypeID"].ToString() + "')";
                        string strinsertCOMPSchedule = "Insert into BM_COMPSchedule (JBM_AutoID,AutoArtID) values('" + clsCustSave.Rows[0]["JBM_AutoID"].ToString() + "','" + clsCustSave.Rows[0]["AutoArtID"].ToString() + "')";
                        SqlCommand cmdstrinsertchapterinfo = new SqlCommand(strinsertchapterinfo, con);
                        cmdstrinsertchapterinfo.ExecuteNonQuery();
                        SqlCommand cmdstrinsertProdInfo = new SqlCommand(strinsertProdInfo, con);
                        cmdstrinsertProdInfo.ExecuteNonQuery();
                        SqlCommand cmdstrinsertStageinfo = new SqlCommand(strinsertStageinfo, con);
                        cmdstrinsertStageinfo.ExecuteNonQuery();
                        SqlCommand cmdstrinsertCOMPSchedule = new SqlCommand(strinsertCOMPSchedule, con);
                        cmdstrinsertCOMPSchedule.ExecuteNonQuery();

                        string fpathID = GetCurrentPageName();
                        if (fpathID == "")
                            return;
                        string strinsertprodAccess = "Insert into " + Init_Tables.gTblProdAccess + "(EmpAutoID,CustAcc,AccPage,AccTime,Process,AutoArtID,JBM_AutoID,Descript) values('" + Session["EmpAutoId"].ToString() + "','" + Session["sCustAcc"].ToString() + "','" + fpathID + "','" + DateTime.Now.ToString("MM/dd/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + "','0','" + strAutoArtID + "', '" + strJAutoID + "','" + strChapterID.Trim() + " added for " + strStage + "')";
                        SqlCommand cmdstrinsertprodAccess = new SqlCommand(strinsertprodAccess, con);
                        cmdstrinsertprodAccess.ExecuteNonQuery();


                        //' Create Folders
                        //string strCustName = Request.QueryString("CN");
                        //if (blnCreateDirectory && strJBM_Flow) 
                        // clsInit.Proc_Create_Directory_New(strCustAccess, lbl_Error, strIntrnlJid, "Vol00000", strIntnlID, strCustName, strChapterID)

                    }
                    else
                    {
                        DtStageInfo = DBProc.GetResultasDataTbl("Select AutoArtID from " + Session["sCustAcc"].ToString() + "_Stageinfo where AutoArtID='" + strAutoArtID + "' and RevFinStage like '" + strStage + "%'", Session["sConnSiteDB"].ToString());
                        if (DtStageInfo.Rows.Count == 0)
                        {
                            string strinsertStageinfo = "Insert into " + Session["sCustAcc"].ToString() + "_Stageinfo(AutoArtId, RevFinStage, ArtStageTypeID, Rev_wf, ReceivedDate, DueDate, Loginby, LoginEnteredDate, LoginRecDate, Ceparallel, XMLParallel,CustomerRecDate,CustomerDueDate,GraphArtStageTypeID) values ('" + strAutoArtID + "', '" + clsCustSave.Rows[0]["FpRevPtrInfo"].ToString() + "', '" + clsCustSave.Rows[0]["ArtStageTypeID"].ToString() + "', '" + clsCustSave.Rows[0]["Rev_wf"].ToString() + "', '" + clsCustSave.Rows[0]["RecDate"].ToString() + "', '" +clsCustSave.Rows[0]["DueDate"].ToString() + "', '" + strUserID + "', '" + clsCustSave.Rows[0]["LoginEnteredDate"].ToString() + "', '" +clsCustSave.Rows[0]["LoginRecDate"].ToString() + "','" + clsCustSave.Rows[0]["CeParallel"].ToString() + "','" + clsCustSave.Rows[0]["XMLParallel"].ToString() + "','" + clsCustSave.Rows[0]["customerrecdate"].ToString() + "','" + clsCustSave.Rows[0]["customerduedate"].ToString() + "','" + clsCustSave.Rows[0]["ArtStageTypeID"].ToString() + "')";
                            string strinsertCOMPSchedule = "Insert into BM_COMPSchedule (JBM_AutoID,AutoArtID) values('" + strJAutoID + "','" + strAutoArtID + "')";
                            SqlCommand cmdstrinsertStageinfo = new SqlCommand(strinsertStageinfo, con);
                            cmdstrinsertStageinfo.ExecuteNonQuery();


                            string fpathID = GetCurrentPageName();
                            if (fpathID == "")
                                return;
                            string strinsertprodAccess = "Insert into " + Init_Tables.gTblProdAccess + "(EmpAutoID,CustAcc,AccPage,AccTime,Process,AutoArtID,JBM_AutoID,Descript) values('" + Session["EmpAutoId"].ToString() + "','" + Session["sCustAcc"].ToString() + "','" + fpathID + "','" + DateTime.Now.ToString("MM/dd/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + "','0','" + strAutoArtID + "', '" + strJAutoID + "','" + strChapterID.Trim() + " added for " + strStage + "')";
                            SqlCommand cmdstrinsertprodAccess = new SqlCommand(strinsertprodAccess, con);
                            cmdstrinsertprodAccess.ExecuteNonQuery();

                            SqlCommand cmdstrinsertCOMPSchedule = new SqlCommand(strinsertCOMPSchedule, con);
                            cmdstrinsertCOMPSchedule.ExecuteNonQuery();
                        }
                        else
                        {
                            string strupdateProdInfo = "Update " + Session["sCustAcc"].ToString() + "_ProdInfo set FpRevPtrInfo='" + strStage + "',JNRevisedDate=" + clsCustSave.Rows[0]["JNRevisedDate"].ToString() + ", CWP_Article='" + clsCustSave.Rows[0]["CWP_Article"].ToString() + "', OPS_wf='" + clsCustSave.Rows[0]["OPS_wf"].ToString() + "' where AutoArtId = '" + strAutoArtID + "'";
                            SqlCommand cmdstrupdateProdInfo = new SqlCommand(strupdateProdInfo, con);
                            cmdstrupdateProdInfo.ExecuteNonQuery();

                            string fpathID = GetCurrentPageName();
                            if (fpathID == "")
                                return;
                            string strinsertprodAccess = "Insert into " + Init_Tables.gTblProdAccess + "(EmpAutoID,CustAcc,AccPage,AccTime,Process,AutoArtID,JBM_AutoID,Descript) values('" + Session["EmpAutoId"].ToString() + "','" + Session["sCustAcc"].ToString() + "','" + fpathID + "','" + DateTime.Now.ToString("MM/dd/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + "','1','" + strAutoArtID + "', '" + strJAutoID + "','" + strChapterID.Trim() + " stage modified in ProdInfo to " + strStage + "')";
                            SqlCommand cmdstrinsertprodAccess = new SqlCommand(strinsertprodAccess, con);
                            cmdstrinsertprodAccess.ExecuteNonQuery();

                        }
                    }
                }
            }
            con.Close();
        }
       public string GetCurrentPageName()
        {
            string sPath = System.Web.HttpContext.Current.Request.Url.AbsolutePath;
            System.IO.FileInfo oInfo = new System.IO.FileInfo(sPath);
            string sRet = oInfo.Name;
            if (sRet != "")
            {
                string PageID = "Select ID from " + Init_Tables.gTblJBM_Pages + " where JBM_PageName='" + sRet + "'";
                DataTable dtID = DBProc.GetResultasDataTbl(PageID, Session["sConnSiteDB"].ToString());
                if (dtID.Rows.Count > 0)
                    sRet = dtID.Rows[0][0].ToString();
            }
            return sRet;
        }
      
        public DataTable BindGrid(bool blnClearMessage, string strSearchIDs = " ")
        {
            try
            {
                DataTable dt = new DataTable();
                string strJAutoID = Session["sJBMAutoID"].ToString();
                string strDateQry = "Replace(convert(varchar(11),b.ReceivedDate,101), ' ', '-')  as [Received Date], Replace(Convert(Varchar(11), b.DueDate, 101), ' ', '-') as [Due Date], Replace(convert(varchar(11),b.DispatchDate,101), ' ', '-') as [Dispatch Date]";
                string strStage = "FP";
                string strcurdept = "";
                //if (opt_Stage.SelectedValue == 1)
                //    strStage = "Sample";
                //else if (opt_Stage.SelectedValue == 2)
                //    strDateQry = "Replace(convert(varchar(11),b.CeRecDate,106), ' ', '-')  as [Received Date], Replace(Convert(Varchar(11), b.CeDueDate, 106), ' ', '-') as [Due Date],Replace(convert(varchar(11),b.CeDispDate,106), ' ', '-') as [Dispatch Date]";
                //else if (opt_Stage.SelectedValue == 3)
                //    strDateQry = "Replace(convert(varchar(11),b.GraphicsRecDate,106), ' ', '-')  as [Received Date], Replace(Convert(Varchar(11), b.GraphicsDueDate, 106), ' ', '-') as [Due Date],Replace(convert(varchar(11),b.GraphicsDispDate,106), ' ', '-') as [Dispatch Date]";
                //else if (opt_Stage.SelectedValue == 5)// ''Added for COMP Schedule implementation on Feb 19, 2018
                //{ 
                //    DrpDListStage.Visible = true;
                //    btnAddCOMPSCHED.Visible = true;
                //    lbl_Status.Visible = true;
                //}

                string strfigs = "";
                if (strStage == "FP" || strStage == "Sample")
                    strfigs = "a.NumofFigures";
                else
                    strfigs = "b.CorrFigs";


                string strQuery = "Select a.Batch as [Batch], a.AutoArtid as [AutoArtid], a.ChapterID as [ChapterID], a.ArtTypeID as [Type]," +
                    "a.PlatformID as [Platform],a.Complexity as [Complexity],a.NumofMSP as [MSS],a.castoff as [CastOff],isnull(a.ActualPages,0) as ActualPages," +
                    "" + strDateQry + ", '" + strStage + "' as [Stage] ,c.ArticleTitle as [Chapter Title],a.Authors as [Author Name],a.NumofTables as [Tbl]," +
                    " " + strfigs + " as [Fig],Replace(Convert(Varchar(11), b.CustomerDueDate, 101), ' ', '-') as [CustDueDate], Replace(Convert(Varchar(11)," +
                    " b.CustomerDueDate, 101), ' ', '-') as CEDUE, ' '  as CERevisedDate ,  ' ' as CEActualDate, ' ' as MSToCOMP,  ' ' as MSToCOMPRevisedDate," +
                    " ' ' as MSToCOMPActualDate , ' ' as FirstPages, ' ' as FirstPagesRevisedDate,  ' ' as FirstPagesActualDate ,' ' as Correctionfp, " +
                    "' '  as  CorrectionfpRevisedDate,  '  '  as CorrectionfpPermsDate, ' '  as  CorrectionfpPubDate , ' '  as CorrectionfpPRDate, ' '  as  CorrectionfpAUDate, " +
                    "' ' as MastersetToComp, ' '  as  MastersetToCompRevisedDate, '  '  as MastersetToCompActualDate , ' ' as RevisedPages, '' as  RevisedPagesRevisedDate," +
                    " ' ' as  RevisedPagesActualDate, ' '  as  Correctionrandom, '' as CorrRandomRevisedDate, ' ' as CorrRandomActualDate, ' ' as FinalPages, " +
                    "' ' as  FinalPagesRevisedDate,  ' ' as FinalPagesActualDate, ' ' as Correction, ' ' as CorrectionRevisedDate, ' ' as CorrectionActualDate, " +
                    "' ' as FilesToPrinter, ' ' as FilesToPrinterRevisedDate, ' ' as FilesToPrinterActualDate from " + Session["sCustAcc"].ToString() + "_chapterinfo a, " +
                    "" + Session["sCustAcc"].ToString() + "_Stageinfo b, " + Session["sCustAcc"].ToString() + "_ProdInfo c where a.AutoArtID=b.AutoArtID and a.AutoArtID=c.AutoArtID " +
                    "and a.JBM_AutoID='" + strJAutoID + "'and b.RevFinStage='" + strStage + "'" + strSearchIDs + strcurdept + " Order by ChapterID";
                //if (opt_Stage.SelectedValue == 5)
                //   strQuery = " Select ' '  as [Batch], BCS.AutoArtID  as [AutoArtid], BCI.Chapterid   as [ChapterID] , '0' as [Type], '2' as [Platform], ' ' as CustDueDate, Replace(Convert(Varchar(11), BCS.[CE-Due], 106), ' ', '-')   as CEDUE, Replace(Convert(Varchar(11), BCS.[CE-Revised] , 106), ' ', '-') as CERevisedDate ,  Replace(Convert(Varchar(11), BCS.[CE-Actual] , 106), ' ', '-')  as CEActualDate, Replace(Convert(Varchar(11), BCS.[MSToComp-Due] , 106), ' ', '-')  as MSToCOMP, Replace(Convert(Varchar(11), BCS.[MSToComp-Revised], 106), ' ', '-')  as MSToCOMPRevisedDate, Replace(Convert(Varchar(11), BCS.[MSToComp-Actual] , 106), ' ', '-') as MSToCOMPActualDate , Replace(Convert(Varchar(11), BCS.[FirstPages-Due], 106), ' ', '-')  as FirstPages, '1' as [Complexity], ' ' as [MSS], ' ' as [CastOff], ' ' as ActualPages, ' ' as [Received Date],  ' 'as [Due Date], ' ' as [Dispatch Date], ' ' as [Tbl], ' ' as [Fig],   Replace(Convert(Varchar(11), BCS.[FirstPages-Revised], 106), ' ', '-') as FirstPagesRevisedDate,  Replace(Convert(Varchar(11), BCS.[FirstPages-Actual], 106), ' ', '-')   as FirstPagesActualDate , Replace(Convert(Varchar(11), BCS.[FpCorrections-Due] , 106), ' ', '-')   as Correctionfp, Replace(Convert(Varchar(11), BCS.[FpCorrections-Revised] , 106), ' ', '-')  as  CorrectionfpRevisedDate,  Replace(Convert(Varchar(11), BCS.[FpCorrections-Perms] , 106), ' ', '-') as CorrectionfpPermsDate, Replace(Convert(Varchar(11), BCS.[FpCorrections-Pub], 106), ' ', '-')  as  CorrectionfpPubDate , Replace(Convert(Varchar(11), BCS.[FpCorrections-PR] , 106), ' ' , '-')  as CorrectionfpPRDate, Replace(Convert(Varchar(11),BCS.[FpCorrections-AU] , 106), ' ' , '-')   as  CorrectionfpAUDate, Replace(Convert(Varchar(11), BCS.[MastersetToComp-Due] , 106), ' ' , '-')   as MastersetToComp, Replace(Convert(Varchar(11),BCS.[MastersetToComp-Revised] , 106), ' ' , '-')   as  MastersetToCompRevisedDate, Replace(Convert(Varchar(11),BCS.[MastersetToComp-Actual] , 106), ' ' , '-')   as MastersetToCompActualDate , Replace(Convert(Varchar(11),BCS.[RevisedPages-Due], 106), ' ' , '-')   as RevisedPages, Replace(Convert(Varchar(11),BCS.[RevisedPages-Revised] , 106), ' ', '-') as RevisedPagesRevisedDate, Replace(Convert(Varchar(11), BCS.[RevisedPages-Actual] , 106), ' ', '-')   as  RevisedPagesActualDate, Replace(Convert(Varchar(11), BCS.[CorrectionsRandoms-Due], 106),'' , '-')   as  Correctionrandom, Replace(Convert(Varchar(11),BCS.[CorrectionsRandoms-Revised], 106), ' ' , '-')   as CorrRandomRevisedDate, Replace(Convert(Varchar(11),BCS.[CorrectionsRandoms-Actual], 106), ' ' , '-')   as CorrRandomActualDate, Replace(Convert(Varchar(11),BCS.[FinalPages-Due], 106), ' ' , '-')   as FinalPages, Replace(Convert(Varchar(11),BCS.[FinalPages-Revised], 106), ' ' , '-')   as  FinalPagesRevisedDate,  Replace(Convert(Varchar(11),BCS.[FinalPages-Actual], 106), ' ' , '-')  as FinalPagesActualDate, Replace(Convert(Varchar(11),BCS.[Corrections-Due], 106), ' ' , '-')   as Correction, Replace(Convert(Varchar(11),BCS.[Corrections-Revised] , 106), ' ' , '-')   as CorrectionRevisedDate, Replace(Convert(Varchar(11),BCS.[Corrections-Actual], 106), ' ' , '-')  as CorrectionActualDate, Replace(Convert(Varchar(11),BCS.[FilesToPrinter-Due], 106), ' ' , '-')  as FilesToPrinter, Replace(Convert(Varchar(11),BCS.[FilesToPrinter-Revised], 106), ' ' , '-') as FilesToPrinterRevisedDate, Replace(Convert(Varchar(11),BCS.[FilesToPrinter-Actual], 106), ' ' , '-')   as FilesToPrinterActualDate from BM_COMPschedule  as BCS, BK_ChapterInfo as BCI	where BCI.AutoArtID=BCS.AutoArtID and BCS.JBM_AutoID='" + strJAutoID + "'" + " Order by ChapterID ";

                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString()); //RecordManager.GetRecord_Multiple_All(strQuery, "TempTable");
                                                                                       //grdVwChapters.DataSource = Dt
                                                                                       //grdVwChapters.DataBind()


                //if (grdVwChapters.Rows.Count == 0)
                //    {
                //        btn_Update.Visible = false;
                //        btnXlsExport.Visible = false;
                //        lbl_Error.Text &= "No jobs found";
                //    }
                //    else {
                //        btn_Update.Visible = true;
                //        btnXlsExport.Visible = false;// 'Nirmal'
                //    }
                //if (opt_Stage.SelectedValue == 0 || opt_Stage.SelectedValue == 1) {
                //    grdVwChapters.Columns(4).Visible = True
                //    grdVwChapters.Columns(5).Visible = True
                //    grdVwChapters.Columns(6).Visible = True
                //    grdVwChapters.Columns(11).Visible = True
                //    grdVwChapters.Columns(12).Visible = True
                //    grdVwChapters.Columns(13).Visible = True
                //    grdVwChapters.Columns(14).Visible = True
                //    grdVwChapters.Columns(15).Visible = True

                //    grdVwChapters.Columns(16).Visible = False  ''Added for CeDue
                //    grdVwChapters.Columns(17).Visible = False  ''Added for CeRevised
                //    grdVwChapters.Columns(18).Visible = False  ''Added for CeActual

                //    grdVwChapters.Columns(19).Visible = False  ''Added for MSToCOMP
                //    grdVwChapters.Columns(20).Visible = False  ''Added for MSToCOMPRevised
                //    grdVwChapters.Columns(21).Visible = False  ''Added for MSToCOMPActual

                //    grdVwChapters.Columns(22).Visible = False  ''Added for FirstPages
                //    grdVwChapters.Columns(23).Visible = False  ''Added for FirstPagesRevised
                //    grdVwChapters.Columns(24).Visible = False  ''Added for FirstPagesActual

                //    grdVwChapters.Columns(25).Visible = False  ''Added for Correctionfp
                //    grdVwChapters.Columns(26).Visible = False  ''Added for CorrectionfpRevised
                //    grdVwChapters.Columns(27).Visible = False  ''Added for CorrectionfpPermsDate
                //    grdVwChapters.Columns(28).Visible = False  ''Added for CorrectionfpPubDate
                //    grdVwChapters.Columns(29).Visible = False  ''Added for CorrectionfpPRDate
                //    grdVwChapters.Columns(30).Visible = False  ''Added for CorrectionfpAUDate

                //    grdVwChapters.Columns(31).Visible = False  ''Added for MastersetToComp
                //    grdVwChapters.Columns(32).Visible = False  ''Added for MastersetToCompRevised
                //    grdVwChapters.Columns(33).Visible = False  ''Added for MastersetToCompActual

                //    grdVwChapters.Columns(34).Visible = False  ''Added for RevisedPages
                //    grdVwChapters.Columns(35).Visible = False  ''Added for RevisedPagesRevised
                //    grdVwChapters.Columns(36).Visible = False  ''Added for RevisedPagesActual

                //    grdVwChapters.Columns(37).Visible = False ''Added for Correctionrandom
                //    grdVwChapters.Columns(38).Visible = False ''Added for CorrectionrandomRevised
                //    grdVwChapters.Columns(39).Visible = False ''Added for CorrectionrandomActual

                //    grdVwChapters.Columns(40).Visible = False ''Added for Final Pages
                //    grdVwChapters.Columns(41).Visible = False ''Added for Final PagesRevised
                //    grdVwChapters.Columns(42).Visible = False ''Added for Final PagesActual

                //    grdVwChapters.Columns(43).Visible = False ''Added for Correction
                //    grdVwChapters.Columns(44).Visible = False ''Added for CorrectionRevised
                //    grdVwChapters.Columns(45).Visible = False ''Added for CorrectionActual

                //    grdVwChapters.Columns(46).Visible = False ''Added for FilesToPrinter
                //    grdVwChapters.Columns(47).Visible = False ''Added for FilesToPrinterRevised
                //    grdVwChapters.Columns(48).Visible = False ''Added for FilesToPrinterAcutal


                //    if (SessionHandler.gJwAccItm.Contains("|PMA|") == false && strSiteID == Init_SiteID.gSiteNoida)
                //                                                                                                                                            grdVwChapters.Columns(12).Visible = false;


                //}
                //                else if (opt_Stage.SelectedValue == 5) {
                //                    grdVwChapters.Columns(3).Visible = True ' Chapter id   ' Code modified for COMP Schedule on Feb 21, 2018
                //                    grdVwChapters.Columns(4).Visible = False
                //                    if (opt_Stage.SelectedValue == 1 )
                //                        grdVwChapters.Columns(5).Visible = True
                //                    else
                //                        grdVwChapters.Columns(5).Visible = False


                //                    grdVwChapters.Columns(6).Visible = False
                //                    grdVwChapters.Columns(7).Visible = False  'MS Pages
                //                    grdVwChapters.Columns(8).Visible = False  ' castoff
                //                    grdVwChapters.Columns(9).Visible = False  ' Act Pages
                //                    grdVwChapters.Columns(10).Visible = False
                //                    grdVwChapters.Columns(11).Visible = False
                //                    grdVwChapters.Columns(12).Visible = False
                //                    grdVwChapters.Columns(13).Visible = False
                //                    grdVwChapters.Columns(14).Visible = False
                //                    grdVwChapters.Columns(15).Visible = False

                //                    grdVwChapters.Columns(16).Visible = True  ''Added for CeDue
                //                    grdVwChapters.Columns(17).Visible = True  ''Added for CeRevised
                //                    grdVwChapters.Columns(18).Visible = True  ''Added for CeActual
                //                    grdVwChapters.Columns(19).Visible = True  ''Added for MSToCOMP
                //                    grdVwChapters.Columns(20).Visible = True  ''Added for MSToCOMPRevised
                //                    grdVwChapters.Columns(21).Visible = True  ''Added for MSToCOMPActual
                //                    grdVwChapters.Columns(22).Visible = True  ''Added for FirstPages
                //                    grdVwChapters.Columns(23).Visible = True  ''Added for FirstPagesRevised
                //                    grdVwChapters.Columns(24).Visible = True  ''Added for FirstPagesActual
                //                    grdVwChapters.Columns(25).Visible = True  ''Added for Correctionfp
                //                    grdVwChapters.Columns(26).Visible = True  ''Added for CorrectionfpRevised
                //                    grdVwChapters.Columns(27).Visible = True  ''Added for CorrectionfpPermsDate
                //                    grdVwChapters.Columns(28).Visible = True  ''Added for CorrectionfpPubDate
                //                    grdVwChapters.Columns(29).Visible = True  ''Added for CorrectionfpPRDate
                //                    grdVwChapters.Columns(30).Visible = True  ''Added for CorrectionfpAUDate
                //                    grdVwChapters.Columns(31).Visible = True  ''Added for MastersetToComp
                //                    grdVwChapters.Columns(32).Visible = True  ''Added for MastersetToCompRevised
                //                    grdVwChapters.Columns(33).Visible = True  ''Added for MastersetToCompActual
                //                    grdVwChapters.Columns(34).Visible = True  ''Added for RevisedPages
                //                    grdVwChapters.Columns(35).Visible = True  ''Added for RevisedPagesRevised
                //                    grdVwChapters.Columns(36).Visible = True  ''Added for RevisedPagesActual
                //                    grdVwChapters.Columns(37).Visible = True ''Added for Correctionrandom
                //                    grdVwChapters.Columns(38).Visible = True ''Added for CorrectionrandomRevised
                //                    grdVwChapters.Columns(39).Visible = True ''Added for CorrectionrandomActual
                //                    grdVwChapters.Columns(40).Visible = True ''Added for Final Pages
                //                    grdVwChapters.Columns(41).Visible = True ''Added for Final PagesRevised
                //                    grdVwChapters.Columns(42).Visible = True ''Added for Final PagesActual
                //                    grdVwChapters.Columns(43).Visible = True ''Added for Correction
                //                    grdVwChapters.Columns(44).Visible = True ''Added for CorrectionRevised
                //                    grdVwChapters.Columns(45).Visible = True ''Added for CorrectionActual
                //                    grdVwChapters.Columns(46).Visible = True ''Added for FilesToPrinter
                //                    grdVwChapters.Columns(47).Visible = True ''Added for FilesToPrinterRevised
                //                    grdVwChapters.Columns(48).Visible = True ''Added for FilesToPrinterAcutal

                //}
                //                else if (opt_Stage.SelectedValue == 2 || opt_Stage.SelectedValue == 3) {

                //                    grdVwChapters.Columns(4).Visible = False
                //                    If opt_Stage.SelectedValue = 1 Then
                //                        grdVwChapters.Columns(5).Visible = True
                //                    Else
                //                        grdVwChapters.Columns(5).Visible = False
                //                    End If

                //                    grdVwChapters.Columns(6).Visible = False
                //                    'grdVwChapters.Columns(11).Visible = False
                //                    grdVwChapters.Columns(15).Visible = False
                //                    grdVwChapters.Columns(12).Visible = False
                //                    grdVwChapters.Columns(13).Visible = False
                //                    grdVwChapters.Columns(14).Visible = False

                //                    grdVwChapters.Columns(16).Visible = False  ''Added for CeDue
                //                    grdVwChapters.Columns(17).Visible = False  ''Added for CeRevised
                //                    grdVwChapters.Columns(18).Visible = False  ''Added for CeActual

                //                    grdVwChapters.Columns(19).Visible = False  ''Added for MSToCOMP
                //                    grdVwChapters.Columns(20).Visible = False  ''Added for MSToCOMPRevised
                //                    grdVwChapters.Columns(21).Visible = False  ''Added for MSToCOMPActual

                //                    grdVwChapters.Columns(22).Visible = False  ''Added for FirstPages
                //                    grdVwChapters.Columns(23).Visible = False  ''Added for FirstPagesRevised
                //                    grdVwChapters.Columns(24).Visible = False  ''Added for FirstPagesActual

                //                    grdVwChapters.Columns(25).Visible = False  ''Added for Correctionfp
                //                    grdVwChapters.Columns(26).Visible = False  ''Added for CorrectionfpRevised
                //                    grdVwChapters.Columns(27).Visible = False  ''Added for CorrectionfpPermsDate
                //                    grdVwChapters.Columns(28).Visible = False  ''Added for CorrectionfpPubDate
                //                    grdVwChapters.Columns(29).Visible = False  ''Added for CorrectionfpPRDate
                //                    grdVwChapters.Columns(30).Visible = False  ''Added for CorrectionfpAUDate

                //                    grdVwChapters.Columns(31).Visible = False  ''Added for MastersetToComp
                //                    grdVwChapters.Columns(32).Visible = False  ''Added for MastersetToCompRevised
                //                    grdVwChapters.Columns(33).Visible = False  ''Added for MastersetToCompActual

                //                    grdVwChapters.Columns(34).Visible = False  ''Added for RevisedPages
                //                    grdVwChapters.Columns(35).Visible = False  ''Added for RevisedPagesRevised
                //                    grdVwChapters.Columns(36).Visible = False  ''Added for RevisedPagesActual

                //                    grdVwChapters.Columns(37).Visible = False ''Added for Correctionrandom
                //                    grdVwChapters.Columns(38).Visible = False ''Added for CorrectionrandomRevised
                //                    grdVwChapters.Columns(39).Visible = False ''Added for CorrectionrandomActual

                //                    grdVwChapters.Columns(40).Visible = False ''Added for Final Pages
                //                    grdVwChapters.Columns(41).Visible = False ''Added for Final PagesRevised
                //                    grdVwChapters.Columns(42).Visible = False ''Added for Final PagesActual

                //                    grdVwChapters.Columns(43).Visible = False ''Added for Correction
                //                    grdVwChapters.Columns(44).Visible = False ''Added for CorrectionRevised
                //                    grdVwChapters.Columns(45).Visible = False ''Added for CorrectionActual

                //                    grdVwChapters.Columns(46).Visible = False ''Added for FilesToPrinter
                //                    grdVwChapters.Columns(47).Visible = False ''Added for FilesToPrinterRevised
                //                    grdVwChapters.Columns(48).Visible = False ''Added for FilesToPrinterAcutal



                //               }

                //if (strSiteID == Init_SiteID.gSiteNoida) {
                //    string SamplAppDate = "select top 1 customerrecdate from " & Init_Tables.gTblStageInfo & " s," & Init_Tables.gTblChapterInfo & " c where c.autoartid=s.autoartid and s.customerrecdate is not null and c.JBM_AutoID='" & strJAutoID & "' order by s.customerrecdate desc"
                //    DataTable SApp = RecordManager.GetRecord_Multiple_All(SamplAppDate, "TempTable");
                //}
                //if (SApp.Rows.Count == 0)
                //    txt_CustomerRecdDt.Text = "";
                //else
                //    txt_CustomerRecdDt.Text = Format(CDate(SApp.Rows(0)(0).ToString()), "dd-MMM-yyyy");



                //string qry_servicetype = "select JBM_AutoID, BM_FullService from " + Init_Tables.gTblJrnlInfo + "  where JBM_AutoID='" + strJAutoID + "'";
                //DataTable dtStype = RecordManager.GetRecord_Multiple_All(qry_servicetype, "TempTable");
                //string strServal = "";
                //if (dtStype.Rows.Count > 0)
                //    strServal = dtStype.Rows[0]["BM_FullService"].ToString();


                //if (SessionHandler.gBMAccItm.Contains("|PMA|") = false)
                //    btn_Update.Visible = false;
                //grdVwChapters.Columns(12).Visible = false;


                //    if (strServal != "0" && SessionHandler.gBMAccItm.Contains("|PMA|") == false && SessionHandler.gBMAccItm.Contains("|AMA|") == true) {
                //    btn_Update.Visible = true;
                //    grdVwChapters.Columns(12).Visible = true;
                //   }


                return dt;
                //txt_ReceivedDate.Text = Format(CDate(Date.Today.ToString()), "dd-MMM-yyyy");
            }
            catch (Exception ex)
            {
                DataTable dt = new DataTable();
                //lbl_Error.Text &= Err.Description
                return dt;
            }

        }
        //public void UploadExcel(string strJAutoID)
        //{
        //    string strFileName = "";
        //    string strFilePath = "";
        //    string strJID = "";
        //    if (FileUploadCover1.PostedFile.FileName == null)
        //    {
        //        return;
        //    }
        //    else
        //    {
        //        if (FileUploadCover1.PostedFile.FileName.EndsWith(".xls") == false & FileUploadCover1.PostedFile.FileName.EndsWith(".xlsx") == false)
        //        {
        //            lbl_Error.Text = "Upload only excel format";
        //            return;
        //        }
        //        FileUploadCover1.PostedFile.SaveAs(clsInit.gApplicationsPath + @"ExcelInput\" + FileUploadCover1.FileName);
        //        strFilePath = clsInit.gApplicationsPath + @"ExcelInput\" + FileUploadCover1.FileName;
        //    }
        //    DataTable dt = new DataTable();
        //    string strQuery = "select * from [Chapter$]";
        //    dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

        //    if (dt.Rows.Count == 0)
        //        return;
        //    try
        //    {
        //        for (int dtrow = 0; dtrow <= dt.Rows.Count - 1; dtrow++)
        //        {
        //            string strcolumns = "";
        //            string strcolumnvalues = "";
        //            string sQuery = "";
        //            string strColChapter = string.Empty;
        //            if (IsNumeric(dt.Rows[dtrow][0].ToString()))
        //                strColChapter = "Chp" + string.Format("{0:D2}", System.Convert.ToInt32(dt.Rows[dtrow][0].ToString().Trim()));
        //            else
        //            { strColChapter = dt.Rows[dtrow][0].ToString().Trim(); }
        //            sQuery = "select autoartid from " + Init_Tables.gTblChapterInfo + " where JBM_AutoID = '" + strJAutoID + "' and chapterID = '" + strColChapter + "'";
        //            DataTable dtexist = DBProc.GetResultasDataTbl(sQuery, Session["sConnSiteDB"].ToString());
        //            string strColumStage = "";
        //            if (dtexist.Rows.Count != 0)
        //            {
        //                for (int dtcol = 1; dtcol <= dt.Columns.Count - 1; dtcol++)
        //                {
        //                    if (dt.Rows[dtrow][dtcol].ToString() != "")
        //                    {
        //                        switch (dt.Columns[dtcol].ColumnName)
        //                        {
        //                            case "Mss":
        //                                {
        //                                    strcolumns += "NumofMSP = '" + dt.Rows[dtrow][dtcol].ToString() + "',";
        //                                    break;
        //                                }

        //                            case "Tables":
        //                                {
        //                                    strcolumns += "NumofTables = '" + dt.Rows[dtrow][dtcol].ToString() + "',";
        //                                    break;
        //                                }

        //                            case "Figs":
        //                                {
        //                                    strcolumns += "NumofFigures = '" + dt.Rows[dtrow][dtcol].ToString() + "',";
        //                                    break;
        //                                }

        //                            case "castoff":
        //                                {
        //                                    strcolumns += "castoff = '" + dt.Rows[dtrow][dtcol].ToString() + "',";
        //                                    break;
        //                                }

        //                            default:
        //                                {
        //                                    if (dt.Columns[dtcol].ColumnName.Contains("Date"))
        //                                    {
        //                                        DataTable dtStge = DBProc.GetResultasDataTbl("select * from " + Init_Tables.gTblStageInfo + " where autoartid='" + dtexist.Rows[0]["autoartid"] + "' and revfinstage='fp'", Session["sConnSiteDB"].ToString());
        //                                        if (dtStge.Rows.Count != 0)
        //                                            strColumStage += dt.Columns[dtcol].ColumnName + "= '" + Format((DateTime)dt.Rows[dtrow][dtcol], "dd-MMM-yy") + "',";
        //                                    }
        //                                    else
        //                                    {
        //                                        strcolumns += dt.Columns[dtcol].ColumnName + "='" + dt.Rows[dtrow][dtcol].ToString() + "',";

        //                                    } 
        //                                    break;
        //                                }
        //                        }
        //                    }
        //                }
        //            }
        //            if (strcolumns.Length != 0)
        //            {
        //                strcolumns = strcolumns.Substring(0, strcolumns.Length - 1);
        //                sQuery = "Update " + Init_Tables.gTblChapterInfo + "  Set " + strcolumns + "  Where JBM_AutoId='" + strJAutoID + "' and chapterID = '" + strColChapter + "'";

        //                if (strColumStage.Length != 0)
        //                {
        //                    strColumStage = strColumStage.Substring(0, strColumStage.Length - 1);
        //                    sQuery += ";Update " + Init_Tables.gTblStageInfo + "  Set " + strColumStage + "  Where autoartid='" + dtexist.Rows[0]["autoartid"] + "' and revfinstage = 'fp'";

        //                }
        //                RecordManager.UpdateRecord(sQuery);
        //            }
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        lbl_Error.Text = ex.Message
        //    }
        //    finally
        //    {
        //        lbl_Error.Text = "Excel document uploaded successfully";
        //    }
        //}
        [SessionExpire]
        public ActionResult ImportExcel(List<HttpPostedFileBase> fileUpload)
        {
            try
            {
                string strJAutoID = Session["sJBMAutoID"].ToString();
                DataSet dsrootdirectory = new DataSet();
                DataSet dsDeptFolders = new DataSet();
                dsrootdirectory = DBProc.GetResultasDataSet("select RootID,RootPath from jbm_rootdirectory where RootID=1", Session["sConnSiteDB"].ToString());

                dsDeptFolders = DBProc.GetResultasDataSet("SELECT   CustType, RootID, FolderDir, FolderIndex, CreateDirectory FROM   JBM_DeptFolders where custtype = '" + Session["sCustAcc"].ToString() + "' and FolderIndex = 'F20-PN' and RootID = 1", Session["sConnSiteDB"].ToString());
                string rootdirectory = "";
                string DeptFolders = "";
                if (dsrootdirectory.Tables[0].Rows.Count > 0)
                {
                    rootdirectory = dsrootdirectory.Tables[0].Rows[0]["RootPath"].ToString();
                }
                if (dsDeptFolders.Tables[0].Rows.Count > 0)
                {
                    DeptFolders = dsDeptFolders.Tables[0].Rows[0]["FolderDir"].ToString();
                    DeptFolders = DeptFolders.Replace("###CustID###", Session["CustomerName"].ToString());
                    DeptFolders = DeptFolders.Replace("###JID###", Session["ProjectID"].ToString());
                }
                string rootpath = @"" + rootdirectory + DeptFolders + "ExcelInput";


                if (fileUpload != null)
                {
                    foreach (HttpPostedFileBase file in fileUpload)
                    {

                        if (file != null)
                        {
                            string strExcelFile = file.FileName;
                            string strSaveFile = "";
                            if ((!System.IO.Directory.Exists(rootpath)))
                            {
                                System.IO.Directory.CreateDirectory(rootpath);
                            }
                            clsCollec.SetFolderPermission(rootpath); // To set permission

                            string xlPath = rootpath + "\\" + strExcelFile;
                            string Ext = Path.GetExtension(xlPath);
                            if (Ext != ".xls" && Ext != ".xlsx")// Then ''Or flBrowse.PostedFile.FileName.Contains(".xlsx"))
                            {
                                // return Json("Please upload excel file");
                                return Json(new { dataSch = "Please upload excel file" }, JsonRequestBehavior.AllowGet);
                            }

                            strSaveFile = System.IO.Path.GetFileNameWithoutExtension(xlPath);
                            strSaveFile = strSaveFile + System.DateTime.Now.ToString("_dd_MM_yyyy_hhmm");
                            string SavefilePath = "";
                            try
                            {
                                SavefilePath = rootpath + "\\" + strSaveFile + ".xlsx";
                                if (!System.IO.File.Exists(SavefilePath))
                                {

                                    System.IO.File.WriteAllBytes(SavefilePath, clsCollec.ReadData(file.InputStream));
                                }
                                else
                                {
                                    return Json(new { dataSch = "Already exists" }, JsonRequestBehavior.AllowGet);
                                }
                            }
                            catch (Exception ex)
                            {
                                return Json(new { dataSch = "Error saving file:" + ex.Message }, JsonRequestBehavior.AllowGet);
                            }


                            try
                            {
                                string FileName = Path.GetFileName(rootpath + "\\" + strSaveFile + ".xlsx");
                                string Extension = Path.GetExtension(rootpath + "\\" + strSaveFile + ".xlsx");
                                //string conStr = "";
                                //switch (Extension)
                                //{
                                //    case ".xls":
                                //        {
                                //            conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
                                //            break;
                                //        }
                                //    case ".xlsx":
                                //        {
                                //            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
                                //            break;
                                //        }
                                //}

                                //conStr = string.Format(conStr, rootpath + "\\" + strSaveFile + ".xlsx", "Yes");
                               // DataTable xlDt= new DataTable();
                                //OleDbConnection connExcel = new OleDbConnection(conStr);
                                //OleDbCommand cmdExcel = new OleDbCommand();
                                //OleDbDataAdapter oda = new OleDbDataAdapter();
                                //xlDt = new DataTable();
                                // cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet              
                                //connExcel.Open();
                                //DataTable dtExcelSchema = new DataTable();
                                //dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null/* TODO Change to default(_) if this is not a reference type */);
                                //string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                //connExcel.Close();

                                //Read Data from First Sheet
                                //connExcel.Open();
                                //cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                                //oda.SelectCommand = cmdExcel;
                                //oda.Fill(xlDt);
                                //connExcel.Close();

                                //Open the Excel file in Read Mode using OpenXml.

                               // string path = @"D:\SmartTrack_Changes\From_aparna\1st April\Chapters1617257709892.xlsx";
                                FileStream stream = System.IO.File.Open(SavefilePath, FileMode.Open, FileAccess.Read);
                                var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                                DataTable xlDt = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                }).Tables["Sheet1"];

                                //using (SpreadsheetDocument doc = SpreadsheetDocument.Open(SavefilePath, false))
                                //{
                                //    //Read the first Sheets from Excel file.
                                //    Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();

                                //    //Get the Worksheet instance.
                                //    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

                                //    //Fetch all the rows present in the Worksheet.
                                //    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                                //    //Create a new DataTable.
                                //   // DataTable dt = new DataTable();

                                //    //Loop through the Worksheet rows.
                                //    foreach (Row row in rows)
                                //    {
                                //        //Use the first row to add columns to DataTable
                                //        if (row.RowIndex.Value == 1)
                                //        {
                                //            foreach (Cell cell in row.Descendants<Cell>())
                                //            {
                                //                xlDt.Columns.Add(GetValue(doc, cell));
                                //            }
                                //        }
                                //        else 
                                //        {
                                //            //Add rows to DataTable.
                                //            xlDt.Rows.Add();
                                //            int i = 0;
                                //            foreach (Cell cell in row.Descendants<Cell>())
                                //            {
                                //                xlDt.Rows[xlDt.Rows.Count - 1][i] = GetValue(doc, cell);
                                //                i++;
                                //            }
                                //        }
                                //    }

                                //}



                                if (xlDt.Rows.Count == 0)
                                {
                                    return Json(new { dataSch = "No records found in excel" }, JsonRequestBehavior.AllowGet);
                                }

                                for (int dtrow = 1; dtrow <= xlDt.Rows.Count - 1; dtrow++)
                                {
                                    string strcolumns = "";                                   
                                    string sQuery = "";
                                    string sqryupd = "";
                                    string strColChapter = string.Empty;
                                   
                                    strColChapter = xlDt.Rows[dtrow][2].ToString().Trim();                                   

                                    string strnotbls = "";
                                    string strnofigs = "";
                                    string strnomsp = "";
                                    string strCastoff = "";

                                    strnotbls = xlDt.Rows[dtrow][15].ToString().Trim();
                                    strnofigs = xlDt.Rows[dtrow][16].ToString().Trim();
                                    strnomsp = xlDt.Rows[dtrow][6].ToString().Trim();
                                    strCastoff = xlDt.Rows[dtrow][7].ToString().Trim();
                                    sQuery = "select autoartid from " + Session["sCustAcc"].ToString() + "_chapterinfo where JBM_AutoID = '" + strJAutoID + "' and chapterID = '" + strColChapter + "'";
                                    DataTable dtexist = DBProc.GetResultasDataTbl(sQuery, Session["sConnSiteDB"].ToString());
                                   
                                    if (dtexist.Rows.Count != 0)
                                    {
                                        strcolumns += "NumofMSP = '" + strnomsp.Trim() + "',";
                                        strcolumns += "NumofFigures = '" + strnofigs + "',";
                                        strcolumns += "NumofTables = '" + strnotbls + "',";
                                        strcolumns += "Castoff = '" + strCastoff + "',";
                                    }
                                    if (strcolumns.Length != 0)
                                    {
                                        SqlConnection con = new SqlConnection();
                                        con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                                        con.Open();
                                        strcolumns = strcolumns.Substring(0, strcolumns.Length - 1);
                                        sqryupd = "Update " + Session["sCustAcc"].ToString() + "_Chapterinfo Set " + strcolumns + "  Where JBM_AutoId='" + strJAutoID + "' and chapterID = '" + strColChapter + "'";
                                        SqlCommand cmd = new SqlCommand(sqryupd, con);
                                        cmd.ExecuteNonQuery();                                      
                                        con.Close();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                return Json(new { dataSch = "Failed. " + ex.Message }, JsonRequestBehavior.AllowGet);
                            }
                        }
                    }
                }
                else
                {
                    return Json(new { dataSch = "Please select the file" }, JsonRequestBehavior.AllowGet);
                }
                //return Json("Excel document uploaded successfully", JsonRequestBehavior.AllowGet);
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]

        public JsonResult UpdateChapters(List<BookInData> BookInDatalst)
        {
            try
            {
                SqlConnection con = new SqlConnection();
                int count = BookInDatalst.Count;
                for (int i = 0; i < count; i++)
                {
                    string BAutoArtID = BookInDatalst[i].BAutoArtid.ToString();
                    string BChapterID = BookInDatalst[i].BChapterID.ToString();
                    string BShort_Stage = BookInDatalst[i].BShort_Stage.ToString();
                    string BType = BookInDatalst[i].BType.ToString();
                    string BPlatform = BookInDatalst[i].BPlatform.ToString();
                    string BComplexity = BookInDatalst[i].BComplexity.ToString();
                    string BMSS = BookInDatalst[i].BMSS.ToString();
                    string BCastOff = BookInDatalst[i].BCastOff.ToString();
                    string BActualPages = BookInDatalst[i].BActualPages != null ? "'" + BookInDatalst[i].BActualPages.ToString() + "'" : "''";
                    string BChapterTitle = BookInDatalst[i].BChapterTitle != null ? "'" + BookInDatalst[i].BChapterTitle.ToString() + "'" : "''";
                    string BAuthorName = BookInDatalst[i].BAuthorName != null ? "'" + BookInDatalst[i].BAuthorName.ToString() + "'" : "''";
                    string BTbl = BookInDatalst[i].BTbl != null ? "'" + BookInDatalst[i].BTbl.ToString() + "'" : "''";
                    string BFig = BookInDatalst[i].BFig != null ? "'" + BookInDatalst[i].BFig.ToString() + "'" : "''";
                    string BDueDate = BookInDatalst[i].BDueDate != null ? "'" + BookInDatalst[i].BDueDate + "'" : "null";
                    string BCustDueDate = BookInDatalst[i].BDueDate != null ? "'" + BookInDatalst[i].BCustDueDate + "'" : "null";
                    string BCEDUE = BookInDatalst[i].BDueDate != null ? "'" + BookInDatalst[i].BCEDue + "'" : "null";
                    string BReceivedDate = BookInDatalst[i].BReceivedDate != null ? "'" + BookInDatalst[i].BReceivedDate + "'" : "null";
                    string BDispatchDate = BookInDatalst[i].BDispatchDate != null ? "'" + BookInDatalst[i].BDispatchDate + "'" : "null";

                    con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                    con.Open();
                    if (BShort_Stage.Trim() == "FP")
                    {
                        SqlCommand cmdchapterinfo = new SqlCommand("UPDATE a SET a.ArtTypeID = " + BType + ", a.PlatformID  = " + BPlatform + ", a.Complexity = " + BComplexity + ", a.NumofMSP = " + BMSS + ", a.castoff = " + BCastOff + ", a.ActualPages = " + BActualPages + ", a.Authors = " + BAuthorName + ", a.NumofTables = " + BTbl + ", a.NumofFigures = " + BFig + " from " + Session["sCustAcc"].ToString() + "_chapterinfo a, " + Session["sCustAcc"].ToString() + "_Stageinfo b, " + Session["sCustAcc"].ToString() + "_ProdInfo c where a.AutoArtID = b.AutoArtID and a.AutoArtID = c.AutoArtID and a.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and b.RevFinStage = 'FP' and a.AutoArtID='" + BAutoArtID + "'", con);
                        cmdchapterinfo.ExecuteNonQuery();
                        SqlCommand cmdStageinfo = new SqlCommand("UPDATE b SET b.ReceivedDate= " + BReceivedDate + ", b.DueDate= " + BDueDate + ", DispatchDate= " + BDispatchDate + ", b.CustomerDueDate= " + BCustDueDate + " from  " + Session["sCustAcc"].ToString() + "_chapterinfo a,  " + Session["sCustAcc"].ToString() + "_Stageinfo b,  " + Session["sCustAcc"].ToString() + "_ProdInfo c where a.AutoArtID = b.AutoArtID and a.AutoArtID = c.AutoArtID and a.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and b.RevFinStage = 'FP' and b.AutoArtID='" + BAutoArtID + "'", con);
                        cmdStageinfo.ExecuteNonQuery();
                        SqlCommand cmdProdInfo = new SqlCommand("UPDATE c SET c.ArticleTitle= " + BChapterTitle + " from " + Session["sCustAcc"].ToString() + "_chapterinfo a, " + Session["sCustAcc"].ToString() + "_Stageinfo b, " + Session["sCustAcc"].ToString() + "_ProdInfo c where a.AutoArtID = b.AutoArtID and a.AutoArtID = c.AutoArtID and a.JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "' and b.RevFinStage = 'FP' and c.AutoArtID='" + BAutoArtID + "'", con);
                        cmdProdInfo.ExecuteNonQuery();
                    }
                    else
                    {
                        //SqlCommand cmd = new SqlCommand(" update BK_Stageinfo set ReceivedDate =null, DueDate=null,RevisedDate=null  where AutoArtID in (select AutoArtID from BK_ChapterInfo  where ChapterID = 'Chp01' and JBM_AutoID = 'BK510') and RevFinStage = 'FP'", con);
                        //SqlCommand cmd = new SqlCommand("update " + Session["sCustAcc"].ToString() + "_Stageinfo set DispatchDate =" + SReceivedDate + ", DueDate=" + SDueDate + ", RevisedDate=" + SRevisedDate + "  where AutoArtID in (select AutoArtID from " + Session["sCustAcc"].ToString() + "_ChapterInfo  where ChapterID = '" + SChapterID + "' and JBM_AutoID = '" + Session["sJBMAutoID"].ToString() + "') and RevFinStage = '" + SShort_Stage + "'", con);
                        //cmd.ExecuteNonQuery();
                    }
                }
                con.Close();
                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
		
		[SessionExpire]
        public ActionResult GetFreelancerCEList(string sCustomer)
        {

            try
            {
                if (sCustomer != "" || sCustomer != null)
                {
                     string strQueryFinal = "SELECT distinct EmpName from JBM_EmployeeMaster where DeptCode='40' and  eType ='OnShoreCE'";

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
        
    }
    public class vSchedule
    {
        public string SChapterID;
        public string SShort_Stage;
        public string SDueDate;
        public string SRevisedDate;
        public string SReceivedDate;
        public string SPubDate;
        public string SPRDate;
        public string SAUDate;
    }
}