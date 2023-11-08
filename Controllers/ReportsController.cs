using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using SmartTrack.Helper;
using System.Management;
using System.DirectoryServices;
using Microsoft.VisualBasic;
using System.IO;
using System.Xml;
using System.Diagnostics;
using RL = ReferenceLibrary;
using Renci.SshNet;

namespace SmartTrack.Controllers
{
    [SessionExpire]
    public class ReportsController : Controller
    {
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        clsCollection clsCollec = new clsCollection();

        // GET: Reports
        public ActionResult CompReport()
        {
            Session["returnURL"] = Request.Url.AbsoluteUri.ToString();
            ViewBag.ReportList = getReportList();
            return View();
        }

        public ActionResult CompReportNew()
        {
            Session["returnURL"] = Request.Url.AbsoluteUri.ToString();
            ViewBag.ReportList = getReportList();
            return View();
        }

        public ActionResult btnReportData_Click(string Reporttype,string ReporttypeText, string Stage, string Query, string DateRangeFrom, string DateRangeTo, string TeamID, string SubTeamID, string CustID, string SearchID,string[] JournalList)
        {
            try
            {
                string strCustomerIDs = "";
                string strAutoid = "";
                string strArtIdFilter = "";
                string strStageFilter = "";
                string strRecFromDate = "";
                string strRecToDate = "";

                DataTable dtWip = new DataTable();
                string strRevStage = Stage;
                string strQueryType = Query;
                string strStageType = clsInit.GetStageType(Reporttype);
                string strCustAcc = Session["sCustAcc"].ToString();
                string tempJID = "";

                if (!string.IsNullOrEmpty(strStageType))
                {
                    strStageType = strStageType.Split('|')[0].ToString();
                }

                if (!Regex.IsMatch(Reporttype, "^(CER|CEFR|Fin|Iss|Fbl|Onl|IssueWIP)$") & (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupJrnls001 | Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002))
                {
                    if (SearchID != "")
                    {
                        var strArtTmp = SearchID.Split(',');
                        for (int i = 0; i < strArtTmp.Length; i++)
                        {
                            if (i != strArtTmp.Length - 1)
                            {
                                strAutoid += "'" + strArtTmp[i].Trim() + "', ";
                            }
                            else
                            {
                                strAutoid += "'" + strArtTmp[i].Trim() + "'";
                            }
                        }
                    }
                }
                
                try
                {
                    clsINI obj = new clsINI();
                    
                    if (JournalList != null)
                    {                        
                        foreach (string strJID in JournalList)
                        {
                            if (tempJID != "")
                            {
                                tempJID = tempJID + ",";
                            }
                            tempJID = tempJID + obj.getJournlIDByTeam(strCustAcc, strJID, Session["sConnSiteDB"].ToString());
                        }
                        strCustomerIDs = " and JBI.JBM_AutoID in (" + tempJID + ")";
                    }
                    else if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupJrnls001 | Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
                    {
                        strCustomerIDs = " and JBI.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' ";
                    }

                    if (TeamID != "" & TeamID != "0" & TeamID != null)
                    {
                        strCustomerIDs = strCustomerIDs + " and JBI.JBM_TeamID in ('" + TeamID + "') ";
                    }

                    if (SubTeamID != "" & SubTeamID != "0" & TeamID != null)
                    {
                        strCustomerIDs = strCustomerIDs + " and JBI.JBM_SubTeam in ('" + SubTeamID + "') ";
                    }

                    if (CustID != "" & CustID != null)
                    {
                        strCustomerIDs = strCustomerIDs + " and JBI.CustID in ('" + CustID + "')";
                    }

                    if (!string.IsNullOrEmpty(strAutoid))
                    {
                        strArtIdFilter = " and (a.AutoArtid in (" + strAutoid + ") or a.IntrnlID in (" + strAutoid + " ) or a.ChapterID in (" + strAutoid + ") or a.DOI in (" + strAutoid + " ))";
                    }
                }
                catch (Exception ex)
                {
                    strCustomerIDs = "";
                }

                // Filter with AutoArtID, InternalID, ChapterID/Article ID, DOI           

                string FproofRecDisp = string.Empty;
                string dispDatestr = string.Empty;
                string ReportType = Reporttype;
                string strWhereCondition = string.Empty;
                string strDispHeader = " IsNull(NULLIF(convert(varchar(12),c.DispatchDate,106),''),'') as [Dispatch Date], ";
                string strOrderbyClause = string.Empty;
                SqlCommand sqlcmdWip;
                var myReader = new SqlDataAdapter();
                var myConnection = DBProc.getConnection(Session["sConnSiteDB"].ToString());//new SqlConnection(Session["sConnection"].ToString());
                var myReaderwip = new SqlDataAdapter();
                var myReader1 = new SqlDataAdapter();
                var ds = new DataSet();
                DataSet dswip = new DataSet();
                string Teaminfo = clsINI.GetProfileString("SQL_Query_Stage", "Team", Server.MapPath("~/Query/SQLQueries.qry"));
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###TEAMINF###", Teaminfo).Replace("###CustAccess###", strCustAcc);
                string strQuerySP = string.Empty;

                string standaloneReport = clsINI.GetProfileString("SQL_Query_Stage", "standaloneReport", Server.MapPath("~/Query/SQLQueries.qry"));

                if (Session["sCustAcc"].ToString() == "TF")
                {
                    strDispHeader = " IsNull(NULLIF(convert(varchar(17),c.DispatchDate,113),''),'') as [Dispatch Date], ";
                }

                if (strQryText == "")
                {
                    return Json(new { result = "Error", data = "Report not configured. Please contact S/W Team" }, JsonRequestBehavior.AllowGet);
                }
                else if (!Regex.IsMatch(ReportType, "^("+ standaloneReport + ")$"))
                {
                    if (Query.ToUpper() == "W")
                    {
                        strDispHeader = "  ";
                        strWhereCondition = " and c.DispatchDate is null ";
                    }
                    else if (Query == "RecDate")
                    {
                        strWhereCondition = " and c.LoginEnteredDate is not null ";
                    }
                    else if (Query == "DueDate")
                    {
                        strWhereCondition = " and c.DueDate is not null ";
                    }
                    else if (Query == "DispDate")
                    {
                        strWhereCondition = " and c.DispatchDate is not null ";
                    }

                    strQryText = strQryText.Replace("###DISPATCH_Header###", strDispHeader);

                    if (Regex.IsMatch(ReportType, "^(WIPAll|RFT)$"))
                    {
                        strWhereCondition = "";
                    }

                    if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002 | Session["sCustAcc"].ToString() == "SA")
                    {
                        if (!Regex.IsMatch(ReportType, "^(WIPAll|RFT)$"))
                        {
                            strQryText = strQryText.Replace("###CustSN_Header###", "(Select CustSN from JBM_CustomerMaster WHERE CustID=JBI.CustID) as [Cust SN], ");
                        }
                        else
                        {
                            strQryText = strQryText.Replace("###CustSN_Header###", " cm.CustSN as [Cust SN], ");
                        }
                    }
                    else
                    {
                        strQryText = strQryText.Replace("###CustSN_Header###", "");
                    }

                    if (DateRangeFrom != "" | DateRangeTo != "")
                    {
                        string strDate1 = "";
                        string strDate2 = "";
                        if (Query.ToUpper() == "W" | strQueryType == "RecDate")
                        {
                            strWhereCondition += " and c.LoginEnteredDate ";
                            strDate1 = " [Rev Login Date] ";
                            strDate2 = " [ePub Login Date] ";
                        }
                        else if (strQueryType == "DueDate")
                        {
                            strWhereCondition += " and c.DueDate ";
                            strDate1 = " [Rev Due Date] ";
                            strDate2 = " [ePub Due Date] ";
                        }
                        else if (strQueryType == "DispDate")
                        {
                            strWhereCondition += " and c.DispatchDate ";
                            strDate1 = " [Rev Dispatch Date] ";
                            strDate2 = " [ePub Dispatch Date] ";
                        }

                        strRecFromDate = DateRangeFrom;
                        strRecToDate = DateRangeTo;
                        if (strRecFromDate != "" & strRecToDate != "")
                        {
                            strWhereCondition += " between cast(convert(varchar(50), '" + strRecFromDate + "',103) as datetime) and cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";

                            strDate1 += " between cast(convert(varchar(50), '" + strRecFromDate + "',103) as datetime) and cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";
                            strDate2 += " between cast(convert(varchar(50), '" + strRecFromDate + "',103) as datetime) and cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";
                        }
                        else if (strRecFromDate != "")
                        {
                            strWhereCondition += " >= cast(convert(varchar(50),'" + strRecFromDate + "',103) as datetime)";

                            strDate1 += " >= cast(convert(varchar(50),'" + strRecFromDate + "',103) as datetime)";
                            strDate2 += " >= cast(convert(varchar(50),'" + strRecFromDate + "',103) as datetime)";
                        }
                        else if (strRecToDate != "")
                        {
                            strWhereCondition += " >= cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";

                            strDate1 += " <= cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";
                            strDate2 += " <= cast(convert(varchar(50),'" + strRecToDate + "',103) as datetime)";
                        }

                        if (Regex.IsMatch(ReportType, "^(RFT)$"))
                        {
                            strWhereCondition = " and (" + strDate1 + " or " + strDate2 + ") ";
                        }
                    }

                    if (ReportType == "CER") // ' added on 12-Jun-2018 to display CER report
                    {
                        if (ReporttypeText == "CE")  // Nirmal on31 may2019
                        {
                            strWhereCondition += " and c.RevfinStage like 'FP%'";
                        }
                        else
                        {
                            strWhereCondition += " and c.RevfinStage like 'CER%'";
                        }
                    }
                    else if (!Regex.IsMatch(ReportType, "^(WIP|IssueWIP|WIPAll)$"))
                    {
                        if (ReportType == "ROF")
                        {
                            strWhereCondition += strCustomerIDs;
                        }
                        else if (ReportType == "PapA")
                        {
                            strWhereCondition += " and c.RevfinStage = '" + ReportType + "'";
                        }
                        else if (ReportType == "PapB")
                        {
                            strWhereCondition += " and c.RevfinStage = '" + ReportType + "'";
                        }
                        else if (ReportType == "RFT")
                        {

                        }
                        else
                        {
                            strWhereCondition += " and c.RevfinStage like '" + ReportType + "%'";
                        }
                    }
                    else
                    {
                        strWhereCondition += "";
                        strWhereCondition += strCustomerIDs;
                    } // ' added on 11-Jun-2018 for Journal filter

                    if (Regex.IsMatch(ReportType, "^(WIPAll|RFT)$"))
                    {
                        strOrderbyClause = "";
                    }
                    else if (ReportType != "WIP" & ReportType != "ROF" & ReportType != "IssueWIP") // ' added on 11-Jun-2018 for Order by Clause
                    {
                        strOrderbyClause = " Order by JBI.Priority desc, c.LoginEnteredDate ";
                    }
                    else
                    {
                        strOrderbyClause = " Order by JBI.Priority desc, [Login Date]";
                    }

                    if (Regex.IsMatch(ReportType, "^(WIP|IssueWIP)$"))
                    {
                        strQryText = strQryText.Replace("###WhereCondition###", strWhereCondition + strCustomerIDs + strArtIdFilter + strStageFilter);
                        strQuerySP = strQryText + strOrderbyClause;

                        if ((!string.IsNullOrEmpty(strArtIdFilter) | !string.IsNullOrEmpty(strCustomerIDs)) && ReportType != "IssueWIP")
                        {
                            if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
                            {
                                if (!string.IsNullOrEmpty(strArtIdFilter) & strQuerySP.IndexOf("Union") > -1)
                                {
                                    strQuerySP = strQuerySP.Substring(0, strQuerySP.IndexOf("Union") - 1);
                                }
                            }
                            else
                            {
                                if (strQuerySP.IndexOf("Union") > -1)
                                {
                                    strQuerySP = strQuerySP.Substring(0, strQuerySP.IndexOf("Union") - 1);
                                    strQuerySP = strQuerySP + strCustomerIDs + strOrderbyClause;
                                }
                            }
                        }
                    }
                    else if (ReportType == "WIPAll")
                    {
                        strQryText = strQryText.Replace("###WhereCondition###", strWhereCondition + strArtIdFilter + strStageFilter);
                        strQuerySP = strQryText;
                    }
                    else if (ReportType == "RFT")
                    {
                        strQryText = strQryText.Replace("###WhereCondition###", strWhereCondition);
                        strQryText = strQryText.Replace("###InnerWhereCondition###", strCustomerIDs);
                        strQuerySP = strQryText;
                    }
                    else
                    {
                        strQryText = strQryText.Replace("###WhereCondition###", strWhereCondition);
                        strQuerySP = strQryText + strCustomerIDs + strArtIdFilter + strStageFilter + strOrderbyClause;

                        if (Query.ToUpper() != "W")
                        {
                            strQuerySP = strQuerySP.Replace("and c.DispatchDate is null", "");
                        }
                    }

                    if (ReportType == "Iss")
                    {
                        strQuerySP = strQuerySP.Replace("###Billed_Table###", ""); // have to update billed qry
                    }

                    if (ReportType == "CER")
                    {
                        strQuerySP = strQuerySP.Replace("c.LoginEnteredDate", "c.CeRecDate");
                        strQuerySP = strQuerySP.Replace("c.DueDate", "c.CeDueDate");
                        strQuerySP = strQuerySP.Replace("c.DispatchDate", "c.CeDispDate");
                    }

                    if (ReporttypeText == "CE")
                    {
                        strQuerySP = strQuerySP.Replace("'CER%'", "'FP%'");
                    }
                }
                else
                {
                    strQuerySP = strQryText;
                    strQuerySP = strQuerySP.Replace("###CustAccess###", strCustAcc);
                    strQuerySP = strQuerySP.Replace("###CustID###", string.IsNullOrEmpty(CustID) ? "" : CustID);
                    strQuerySP = strQuerySP.Replace("###JID###", tempJID.Replace("'", "''"));
                    strQuerySP = strQuerySP.Replace("###TeamID###", string.IsNullOrEmpty(TeamID) ? "" : TeamID);
                    strQuerySP = strQuerySP.Replace("###CustGroup###", Session["sCustGroup"].ToString());
                    strQuerySP = strQuerySP.Replace("###DateField###", Query);
                    strQuerySP = strQuerySP.Replace("###FromDate###", DateRangeFrom);
                    strQuerySP = strQuerySP.Replace("###ToDate###", DateRangeTo);
                    strQuerySP = strQuerySP.Replace("###ArticleID###", string.IsNullOrEmpty(strAutoid) ? "" : strAutoid);
                }

                if (strCustAcc == "BK" | strCustAcc == "MG")
                {
                    strQuerySP = strQuerySP.Replace("_ArticleInfo", "_ChapterInfo");
                    strQuerySP = Regex.Replace(strQuerySP, "_ArticleInfo", "_ChapterInfo", RegexOptions.IgnoreCase);
                }

                try
                {
                    dtWip = new DataTable();
                    sqlcmdWip = new SqlCommand(strQuerySP, myConnection);
                    sqlcmdWip.CommandType = CommandType.Text;
                    sqlcmdWip.CommandTimeout = 600;
                    sqlcmdWip.Connection = myConnection;
                    myReaderwip = new SqlDataAdapter(sqlcmdWip);
                    myReaderwip.Fill(dtWip);
                    myReader.Dispose();
                    myReaderwip.Dispose();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
                }



                if (dtWip.Rows.Count != 0 & dtWip.Columns.Contains("ArtstageTypeID"))
                {
                    DataTable dtProArtStaDes = DBProc.GetResultasDataTbl("Select DeptActivity,StageDesc,Stage from " + Init_Tables.gTblProdArtStatusDesDept + "", Session["sConnSiteDB"].ToString());
                    DataTable dtEmpDet = DBProc.GetResultasDataTbl("Select EmpLogin,EmpAutoID,EmpName,EmpSurname from " + Init_Tables.gTblEmployee + " ", Session["sConnSiteDB"].ToString());

                    string CurrStage = string.Empty;
                    int errorRow = 0;

                    for (int i = 0, loopTo1 = dtWip.Rows.Count - 1; i <= loopTo1; i++)
                    {
                        errorRow = i;
                        string tmpFromActivity = "";
                        string tmpCurActivity = "";
                        string tmpToActivity = "";

                        if (dtWip.Rows[i]["ArtstageTypeID"].ToString() != "")
                        {
                            try
                            {
                                CurrStage = "";

                                if (dtWip.Columns.Contains("Current Stage"))
                                {
                                    CurrStage = fn_CurrentstageFull(dtWip.Rows[i]["ArtstageTypeID"].ToString(), dtProArtStaDes, dtEmpDet);
                                    dtWip.Rows[i]["Current Stage"] = CurrStage;
                                }

                                if (dtWip.Columns.Contains("Current Activity"))
                                {
                                    if (dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(0, 2) == "_S")
                                    {
                                        dtWip.Rows[i]["Current Activity"] = CurrentActivity(dtProArtStaDes.Select("DeptActivity = '" + dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(7, 5) + "' and DeptActivity is not Null")[0][1].ToString(), dtWip.Rows[i]["pagcorr"].ToString(), dtWip.Rows[i]["TeCorr"].ToString());
                                    }
                                    else if (CurrStage.Contains("_Hold")) // ' Code added on 22-Aug-2018 as per Mr. Siva instruction
                                    {
                                        dtWip.Rows[i]["Current Activity"] = "Hold";
                                        tmpFromActivity = dtProArtStaDes.Select("DeptActivity = '" + dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(0, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                        tmpCurActivity = "Hold";
                                    }
                                    else
                                    {
                                        string strCurActivity = "";
                                        string strFromActivity = "";
                                        string strToActivity = "";
                                        try
                                        {
                                            strFromActivity = dtProArtStaDes.Select("DeptActivity = '" + dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(0, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                            strToActivity = dtProArtStaDes.Select("DeptActivity = '" + dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(12, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                            strFromActivity = IIf((strFromActivity ?? "") == (strToActivity ?? ""), "", strFromActivity).ToString();
                                            tmpFromActivity = strFromActivity;
                                            tmpToActivity = strToActivity;
                                            strFromActivity = IIf(string.IsNullOrEmpty(strFromActivity), "", " (" + strFromActivity + ")").ToString();
                                        }
                                        catch (Exception ex)
                                        {
                                            strFromActivity = "";
                                        }

                                        strCurActivity = CurrentActivity(dtProArtStaDes.Select("DeptActivity = '" + dtWip.Rows[i]["ArtstageTypeID"].ToString().Substring(12, 5) + "' and DeptActivity is not Null")[0][1].ToString(), dtWip.Rows[i]["pagcorr"].ToString(), dtWip.Rows[i]["TeCorr"].ToString());
                                        if (strCurActivity.ToLower().Contains("corr"))
                                        {
                                            strCurActivity = strCurActivity + strFromActivity;
                                        }

                                        tmpCurActivity = strCurActivity;

                                        dtWip.Rows[i]["Current Activity"] = strCurActivity;

                                    }

                                    if (dtWip.Columns.Contains("WorkFlowDesc"))
                                    {
                                        List<string> WFList = dtWip.Rows[i]["WorkFlowDesc"].ToString().Split(',').ToList();
                                        string WF = dtWip.Rows[i]["WorkFlowDesc"].ToString();
                                        int pos = -1;

                                        if (tmpCurActivity == "QC Check")
                                        {
                                            tmpCurActivity = "QC";
                                        }

                                        if (!WFList.Contains(tmpCurActivity))
                                        {                                        
                                            if (WFList.Contains(tmpFromActivity))
                                            {
                                                pos = WFList.IndexOf(tmpFromActivity) + 1;
                                            }
                                            else if (WFList.Contains(tmpToActivity))
                                            {
                                                pos = WFList.IndexOf(tmpToActivity);
                                            }
                                            else if (tmpCurActivity.StartsWith("Auto-"))
                                            {
                                                //string autoProcess = tmpCurActivity.Substring(5, tmpCurActivity.Length);
                                                string autoProcess = tmpCurActivity.Substring(5);
                                                autoProcess = autoProcess == "Pagination" ? "Pag" : autoProcess;

                                                if (WFList.Contains(autoProcess))
                                                {
                                                    pos = WFList.IndexOf(autoProcess);
                                                }
                                                else
                                                {
                                                    pos = 1;
                                                }
                                            }
                                            else
                                            {
                                                pos = 1;
                                            }

                                            if (pos != -1)
                                            {
                                                WFList.Insert(pos, tmpCurActivity);
                                                dtWip.Rows[i]["WorkFlowDesc"] = string.Join(",", WFList);
                                            }
                                        }
                                        else if (tmpFromActivity != "" & !WFList.Contains(tmpFromActivity))
                                        {
                                            if (WFList.Contains(tmpToActivity))
                                            {
                                                pos = WFList.IndexOf(tmpToActivity);
                                            }

                                            if (pos != -1)
                                            {
                                                WFList.Insert(pos, tmpFromActivity);
                                                dtWip.Rows[i]["WorkFlowDesc"] = string.Join(",", WFList);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (dtWip.Columns.Contains("Current Stage"))
                                {
                                    dtWip.Rows[errorRow]["Current Stage"] = "";
                                }

                                if (dtWip.Columns.Contains("Current Activity"))
                                {
                                    dtWip.Rows[errorRow]["Current Activity"] = "";
                                }
                            }

                        }
                        else
                        {
                            if (dtWip.Columns.Contains("Current Stage"))
                            {
                                dtWip.Rows[errorRow]["Current Stage"] = "";
                            }

                            if (dtWip.Columns.Contains("Current Activity"))
                            {
                                dtWip.Rows[errorRow]["Current Activity"] = "";
                            }
                        }
                    }

                }
                
                List<string> activityOrder = clsINI.GetProfileString("SQL_Query_Stage", "activityOrder", Server.MapPath("~/Query/SQLQueries.qry")).Split(',').ToList();
                List<string> stageOrder = clsINI.GetProfileString("SQL_Query_Stage", "stageOrder", Server.MapPath("~/Query/SQLQueries.qry")).Split(',').ToList();

                List<string[]> incomingJob = new List<string[]>();
                if (dtWip.Columns.Contains("StageName") & dtWip.Columns.Contains("Login Date"))
                {
                    incomingJob = dtWip.AsEnumerable().GroupBy(g => new { Stage = g["StageName"], LoginDate = g["Login Date"] }).Select(s => new string[] { s.Key.Stage.ToString(), s.Key.LoginDate.ToString(), s.Count().ToString() }).OrderBy(o => stageOrder.IndexOf(o[0].ToString().ToUpper())).ToList();
                }

                List<string[]> overdue = new List<string[]>();
                if (dtWip.Columns.Contains("StageName") & dtWip.Columns.Contains("Due Date") & dtWip.Columns.Contains("Aged") & dtWip.Columns.Contains("Current Activity"))
                {
                    overdue = dtWip.AsEnumerable().Where(w => w["Aged"].ToString().Equals("1") & w["Current Activity"].ToString() !="Hold").GroupBy(g => new { Stage = g["StageName"], LoginDate = g["Due Date"] }).Select(s => new string[] { s.Key.Stage.ToString(), s.Key.LoginDate.ToString(), s.Count().ToString() }).OrderBy(o => stageOrder.IndexOf(o[0].ToString().ToUpper())).ToList();
                }

                List<object[]> activity = new List<object[]>();
                if (dtWip.Columns.Contains("Current Activity"))
                {
                    activity = dtWip.AsEnumerable().GroupBy(g => new { Activity = g["Current Activity"] }).Select(s => new[] { s.Key.Activity, s.Count() }).OrderByDescending(o => activityOrder.AsEnumerable().Reverse().ToList().IndexOf(o[0].ToString().ToUpper())).ToList();
                }

                var incomingJobDic = JsonConvert.SerializeObject(changeDateWiseToCategory(incomingJob));
                var overdueDic = JsonConvert.SerializeObject(changeDateWiseToCategory(overdue, -1));
                var activityDic = JsonConvert.SerializeObject(activity);

                var jsonData = (dtWip.Rows.Count > 0) ? JsonConvert.SerializeObject(dtWip) : "";


                var jsonResult = Json(new { result = "Success", data = jsonData, overdue = overdueDic, incomingJob = incomingJobDic, activity = activityDic }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private Dictionary<string, int[]> changeDateWiseToCategory(List<string[]> Job, int addDay = 0)
        {
            Dictionary<string, int[]> dic = new Dictionary<string, int[]>();
            try
            {
                foreach (var c in Job)
                {
                    try
                    {
                        DateTime loginDate = Convert.ToDateTime(c[1]);
                        DateTime currDate = DateTime.Now.Date;
                        string name = c[0].ToString();
                        int count = Convert.ToInt16(c[2]);

                        if (!dic.ContainsKey(name)) { dic.Add(name, new int[] { 0, 0, 0, 0 }); }

                        if (currDate.AddDays(addDay) == loginDate) { dic[name][0] += count; }

                        if (currDate.AddDays(-1 + addDay) == loginDate) { dic[name][1] += count; }

                        //if (loginDate >= currDate.AddDays(-1 + addDay) && loginDate <= currDate.AddDays(-1)) { dic[name][1] += count; }

                        if (loginDate <= currDate.AddDays(addDay) && loginDate > currDate.AddDays(-7 + addDay)) { dic[name][2] += count; }

                        dic[name][3] += count;
                    }
                    catch
                    {
                        continue;
                    }
                }
                return dic;
            }
            catch (Exception ex)
            {
                return new Dictionary<string, int[]>();
            }
        }
        
        public string fn_CurrentstageFull(string ArtStageId, DataTable dtStageDesc, DataTable dtEmpDet)
        {
            string fn_CurrentstageFullRet = "";

            string firstvalExp, secondvalExp, thirdvalExp, empnamevalExp, DatevalExp;
            string empname = "";
            try
            {
                if (ArtStageId.Substring(0, 2) == "_S")
                {
                    string ArtStageIdf = ArtStageId.Substring(0, 6);
                    ArtStageIdf = ArtStageIdf.Replace("_", "");
                    string ArtStageIds = ArtStageId.Substring(7, 5);
                    firstvalExp = dtStageDesc.Select("DeptActivity = '" + ArtStageIdf + "' and DeptActivity is not Null")[0][2].ToString();
                    secondvalExp = dtStageDesc.Select("DeptActivity = '" + ArtStageIds + "' and DeptActivity is not Null")[0][2].ToString();

                    if (!string.IsNullOrEmpty(ArtStageId.Substring(13, 6)))
                    {
                        string empnam = ArtStageId.Substring(13, 7);
                        empnam = empnam.Replace("_", "");
                        if (empnam.IndexOf("_") >= 1)
                        {
                            empname = empnam.Substring(0, empnam.IndexOf("_") - 1);
                        }
                        else
                        {
                            empname = empnam;
                        }
                    }

                    empnamevalExp = dtEmpDet.Select("Emplogin = '" + empname + "' and Emplogin is not Null")[0][2].ToString();
                    DatevalExp = ArtStageId.Substring(24, ArtStageId.Length - 24);
                    fn_CurrentstageFullRet = firstvalExp + "_" + secondvalExp + " by " + empnamevalExp + "_" + DatevalExp;
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("/", "-").Replace("__", "_");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Project Management", "PM");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Issue Management", "IM");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Sivagnanamoorthy M", "System");
                }
                else
                {
                    firstvalExp = dtStageDesc.Select("DeptActivity = '" + ArtStageId.Substring(0, 5) + "' and DeptActivity is not Null")[0][2].ToString();
                    secondvalExp = dtStageDesc.Select("DeptActivity = '" + ArtStageId.Substring(6, 5) + "' and DeptActivity is not Null")[0][2].ToString();
                    thirdvalExp = dtStageDesc.Select("DeptActivity = '" + ArtStageId.Substring(12, 5) + "' and DeptActivity is not Null")[0][2].ToString();
                    if (!string.IsNullOrEmpty(ArtStageId.Substring(18, 6)))
                    {

                        string strGetLSTArt = ArtStageId.Substring(18);
                        string empnam = strGetLSTArt.Substring(0, strGetLSTArt.IndexOf("_"));
                        if (empnam.IndexOf("_") >= 1)
                        {
                            empname = empnam.Substring(0, empnam.IndexOf("_") - 1);
                        }
                        else
                        {
                            empname = empnam;
                        }
                    }

                    empnamevalExp = dtEmpDet.Select("Emplogin = '" + empname + "' and Emplogin is not Null")[0][2].ToString();
                    DatevalExp = ArtStageId.Substring(24, ArtStageId.Length - 24);
                    if (string.IsNullOrEmpty(empnamevalExp))
                    {
                        empnamevalExp = ArtStageId.Substring(18, 5);
                    }

                    if ((secondvalExp ?? "") == (thirdvalExp ?? "") | (firstvalExp ?? "") == (thirdvalExp ?? ""))
                    {
                        fn_CurrentstageFullRet = firstvalExp + "_" + secondvalExp + " by " + empnamevalExp + "_" + DatevalExp;
                    }
                    else
                    {
                        fn_CurrentstageFullRet = firstvalExp + "_" + secondvalExp + " to " + thirdvalExp + " by " + empnamevalExp + "_" + DatevalExp;
                    }

                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("/", "-").Replace("__", "_");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Project Management", "PM");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Issue Management", "IM");
                    fn_CurrentstageFullRet = fn_CurrentstageFullRet.Replace("Sivagnanamoorthy M", "System");
                } // If ArtStageId.Substring(0, 2) Like "_S" Then

                return fn_CurrentstageFullRet;
            }
            catch (Exception ex)
            {
                string ss = ex.Message;
                return fn_CurrentstageFullRet;
            }
        }

        public string CurrentActivity(string ArtStageId, string PageCorr, string TECorr)
        {
            string CurrentActivityRet = "";
            string TempVar = "";
            string TempDept = "";
            string StrValExp = "";
            int IntPage = 0;
            try
            {
                StrValExp = ArtStageId;
                if (StrValExp == "Pag")
                {
                    IntPage = Convert.ToInt32(IIf(string.IsNullOrEmpty(PageCorr), IntPage, PageCorr));
                    CurrentActivityRet = Convert.ToString(IIf(IntPage >= 1, StrValExp + " Corr", StrValExp));
                }
                // CurrentActivity = "Corr"
                else if (StrValExp == "TE")
                {
                    IntPage = Convert.ToInt32(IIf(string.IsNullOrEmpty(TECorr), IntPage, TECorr));
                    CurrentActivityRet = Convert.ToString(IIf(IntPage >= 1, StrValExp + " Check", StrValExp));
                }
                else if (StrValExp == "QC")
                {
                    CurrentActivityRet = StrValExp + " Check";
                }
                else
                {
                    CurrentActivityRet = StrValExp;
                }

                CurrentActivityRet = Convert.ToString(IIf(string.IsNullOrEmpty(CurrentActivityRet), "&nbsp;", CurrentActivityRet));
                CurrentActivityRet = Convert.ToString(IIf(CurrentActivityRet == "Project Management", "PM", CurrentActivityRet));
                CurrentActivityRet = Convert.ToString(IIf(CurrentActivityRet == "Normalization", "ESA", CurrentActivityRet));
            }
            catch (Exception ex)
            {
            }

            return CurrentActivityRet;
        }


        private List<SelectListItem> getReportList(string strMenuItem = "CompRptMenuItem")
        {
            DataTable dt = new DataTable();

            dt = DBProc.GetResultasDataTbl("Select " + strMenuItem + " from " + Init_Tables.gTblAccountTypeDesc + " where CustAccess='" + Session["sCustAcc"].ToString() + "'", Session["sConnSiteDB"].ToString());
            List<SelectListItem> list = new List<SelectListItem>();

            List<string> reportOrder = clsINI.GetProfileString("SQL_Query_Stage", "reportOrder", Server.MapPath("~/Query/SQLQueries.qry")).Split(',').ToList();

            if (dt.Rows.Count != 0)
            {
                string[] ReportLists = dt.Rows[0][strMenuItem].ToString().Split('|');

                foreach (string ReportList in ReportLists)
                {
                    if (ReportList != "")
                    {
                        string[] obj = ReportList.Split('-');
                        SelectListItem selectListItem = new SelectListItem();
                        selectListItem.Text = obj[0].ToString();
                        selectListItem.Value = obj[1].ToString();
                        list.Add(selectListItem);
                    }
                }

            }
            return list.OrderByDescending(o => reportOrder.AsEnumerable().Reverse().ToList().IndexOf(o.Value.ToUpper())).ToList();
        }

        public ActionResult getJobDetails(string strJBMID, string strIss)
        {
            DataTable dt = new DataTable();
            string strEpubDF = Session["sCustAcc"].ToString() == "AI" ? ",EpubDate AS [E-pubDate]" : "";
            dt = DBProc.GetResultasDataTbl("Select AutoArtID As [Job ID],IntrnlID As [Internal ID],ChapterID As [Article ID],iss As [Iss],DOI,NumofMSP As [MSP]" + strEpubDF + ",StartPage As [Start Page],EndPage As [End Page],(case when FinUpload is null then 'N' else 'Y' end) as [Fin Uploaded] from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " where JBM_AutoID='" + strJBMID + "' and ISNULL(iss,'')='" + strIss + "'", Session["sConnSiteDB"].ToString());

            var jsonString = JsonConvert.SerializeObject(dt);
            return Json(jsonString, JsonRequestBehavior.AllowGet);
        }

        object IIf(bool expression, object truePart, object falsePart)
        {
            return expression ? truePart : falsePart;
        }


        public ActionResult SpecialInstruction()
        {
            return View();
        }

        public ActionResult getSpecialInstructionReport(string ReportType)
        {
            try
            {
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString());

                if (strQryText == "")
                {
                    return Json(new { result = "Error", data = "Report not configured. Please contact S/W Team" }, JsonRequestBehavior.AllowGet);
                }

                DataTable dt = new DataTable();

                dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = (dt.Rows.Count > 0) ? JsonConvert.SerializeObject(dt) : "";
                var jsonResult = Json(new { result = "Success", data = jsonString }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult DriveSpace()
        {
            return View();
        }

        public ActionResult getServerDetails()
        {
            try
            {
                //string strQuery = "select distinct '\\\\'+ServerIP As ServerIP from tblApplication";

                //string strQuery = "select distinct '\\\\'+ServerIP+'\\' as ServerIP from tblApplication union select 'C:' ServerIP union select 'D:' ";
                string strQuery = "select distinct ServerIP as ServerIP from tblApplication";

                DataTable dt = new DataTable();

                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                Dictionary<string, string> arr = new Dictionary<string, string>();

                List<Dictionary<string, string>> lst = new List<Dictionary<string, string>>();

                foreach (DataRow dr in dt.Rows)
                {
                    ulong freespace = 0;
                    ulong totalspace = 0;
                    //bool a = clsINI.GetDriveSpace(dr["ServerIP"].ToString(), ref freespace, ref totalspace);
                    CalculateFreeUsed(dr["ServerIP"].ToString(), ref lst);

                    //decimal fs = freespace / 1024 / 1024 / 1024;
                    //decimal ts = totalspace / 1024 / 1024 / 1024;
                    //decimal percentage = 0;

                    //if (freespace != 0 && totalspace != 00)
                    //{
                    //    percentage = Math.Round(((ts - fs) / ts * 100), 2);
                    //}

                    //arr = new Dictionary<string, string>();
                    //arr.Add("Drive", dr["ServerIP"].ToString());
                    //arr.Add("Free Space", fs.ToString() + " GB");
                    //arr.Add("Total Space", ts.ToString() + " GB");
                    //arr.Add("Percentage", percentage.ToString());

                    //lst.Add(arr);
                }

                var jsonString = JsonConvert.SerializeObject(lst);

                var jsonResult = Json(new { result = "Success", data = jsonString }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private void CalculateFreeUsed(string srvname, ref List<Dictionary<string, string>> lst)
        {
            Dictionary<string, string> arr = new Dictionary<string, string>();
            try
            {
                // Connection credentials to the remote computer, not needed if the logged account has access  
                ConnectionOptions oConn = new ConnectionOptions();
                oConn.Username = @"kgl\smarttrack.web";
                oConn.Password = "pass#@321";
                string strNameSpace = @"\\";
                if ((srvname != ""))
                    strNameSpace = (strNameSpace + srvname);
                else
                    strNameSpace = (strNameSpace + ".");
                strNameSpace = (strNameSpace + @"\root\cimv2");
                ManagementScope oMs = new ManagementScope(strNameSpace, oConn);

                // Get Fixed disk state  
                ObjectQuery oQuery = new ObjectQuery("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3");
                // Execute the query  
                ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oMs, oQuery);
                // Get the results  
                ManagementObjectCollection oReturnCollection = oSearcher.Get();

                // loop through found drives and write out info  
                decimal D_Freespace = 0;
                decimal D_Totalspace = 0;
                decimal percentage = 0;

                foreach (ManagementObject oReturn in oReturnCollection)
                {
                    arr = new Dictionary<string, string>();
                    // Disk name        
                    D_Freespace = Convert.ToUInt64(oReturn["FreeSpace"]);
                    D_Totalspace = Convert.ToUInt64(oReturn["Size"]);

                    decimal fs = D_Freespace / 1024 / 1024 / 1024;
                    decimal ts = D_Totalspace / 1024 / 1024 / 1024;

                    if (D_Freespace != 0 && D_Totalspace != 0)
                    {
                        percentage = Math.Round(((D_Totalspace - D_Freespace) / D_Totalspace * 100), 2);
                    }
                    arr.Add("Server", srvname);
                    arr.Add("Disk", oReturn["Name"].ToString());

                    // Free Space in bytes  
                    arr.Add("Free Space", Math.Round(fs, 2) + " GB");

                    // Total Space in bytes  
                    arr.Add("Total Space", Math.Round(ts, 2) + " GB");
                    arr.Add("Percentage", percentage.ToString());

                    lst.Add(arr);

                }

            }
            catch (Exception ex)
            {
                arr.Add("Server", srvname);
                arr.Add("Disk", ex.Message);
                arr.Add("Free Space", "");
                arr.Add("Total Space", "");
                arr.Add("Percentage", "");
                lst.Add(arr);
            }
        }

        public ActionResult ApplicationHandling()
        {
            return View();
        }

        public ActionResult getApplicationDetails()
        {
            try
            {
                string strQuery = "select distinct ServerIP,ExeName,ExePath,'' as [Action] from tblApplication order by ServerIP,ExeName,ExePath";

                DataTable dt = new DataTable();

                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var jsonString = (dt.Rows.Count > 0) ? JsonConvert.SerializeObject(dt) : "";
                var jsonResult = Json(new { result = "Success", data = jsonString }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }


        ManagementScope myScope;
        ConnectionOptions connOptions;
        ManagementObjectSearcher objSearcher;
        ManagementOperationObserver opsObserver;
        ManagementClass manageClass;
        DirectoryEntry entry;
        DirectorySearcher searcher;
        DirectorySearcher userSearcher;
        string[] columnNames = new[] { "Caption", "ComputerName", "Description", "Name", "Priority", "ProcessID", "SessionId" };
        DataColumn[] dc = new DataColumn[7];



        private void ConnectToRemoteMachine(string remoteSystem)
        {
            DataTable dt = new DataTable();

            for (int i = 0; i <= columnNames.Length - 1; i++)
                dc[i] = new DataColumn(columnNames[i], typeof(string));
            dt.Columns.AddRange(dc);

            // Dim remoteSystem As String = "xinchhpdl360-app3"
            //string remoteSystem = "XINCH-MIS";
            // Dim remoteSystem As String = "10.18.11.35"

            string procSearch = "notepad";

            string userName = @"kgl\smarttrack.web";
            string password = "pass#@321";


            string myDomain = "KWGLOBAL.com";

            try
            {
                connOptions = new ConnectionOptions();
                connOptions.Impersonation = ImpersonationLevel.Impersonate;
                connOptions.EnablePrivileges = true;
                connOptions.Authentication = AuthenticationLevel.Default;
                string machineName = remoteSystem;
                if (machineName.ToUpper() == Environment.MachineName.ToUpper())
                    myScope = new ManagementScope(@"\root\cimv2", connOptions);
                else
                {
                    connOptions.Username = userName;
                    connOptions.Password = password;
                    myScope = new ManagementScope((Convert.ToString(@"\\") + machineName) + @"\root\cimv2", connOptions);
                }

                myScope.Connect();
                objSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_Process");
                opsObserver = new ManagementOperationObserver();
                objSearcher.Scope = myScope;
                string[] sep = new[] { Constants.vbLf, Constants.vbTab };

                Console.WriteLine("Authentication sucessful. Getting processes..");
                dt.Rows.Clear();

                foreach (ManagementObject obj in objSearcher.Get())
                {
                    string caption = obj.GetText(TextFormat.Mof);
                    string[] split = caption.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                    DataRow dr = dt.NewRow();
                    // Iterate through the splitter
                    for (int i = 0; i <= split.Length - 1; i++)
                    {
                        if (split[i].Split('=').Length > 1)
                        {
                            string[] procDetails = split[i].Split('=');
                            procDetails[1] = procDetails[1].Replace("\"", "");
                            procDetails[1] = procDetails[1].Replace(';', ' ');
                            switch (procDetails[0].Trim().ToLower())
                            {
                                case "caption":
                                    {
                                        dr[dc[0]] = procDetails[1];
                                        break;
                                    }

                                case "csname":
                                    {
                                        dr[dc[1]] = procDetails[1];
                                        break;
                                    }

                                case "description":
                                    {
                                        dr[dc[2]] = procDetails[1];
                                        break;
                                    }

                                case "name":
                                    {
                                        dr[dc[3]] = procDetails[1];
                                        break;
                                    }

                                case "priority":
                                    {
                                        dr[dc[4]] = procDetails[1];
                                        break;
                                    }

                                case "processid":
                                    {
                                        dr[dc[5]] = procDetails[1];
                                        break;
                                    }

                                case "sessionid":
                                    {
                                        dr[dc[6]] = procDetails[1];
                                        break;
                                    }
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }

                //MsgBox(dt.ToString);
                // bindingSource1.DataSource = dt.DefaultView
                foreach (DataColumn col in dt.Columns)
                {
                }
            }
            // grpStartNewProcess.Enabled = True
            // btnEndProcess.Enabled = True
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message)
                // Console.WriteLine(Information.Err.Description);
            }
            finally
            {
            }
        }

        public void Restart(string remoteSystem)
        {

            // Dim Freebytes As ULong
            // Dim Totalbytes As ULong
            // GetDriveSpace("\\chenas03\cenpro\", Freebytes, Totalbytes)

            // CalculateFreeUsed("XINCH-MIS");
            // ConnectRemoteMachine("XINCH-MIS")

            ConnectToRemoteMachine("XINCH-MIS");
            StartNew("XINCH-MIS");
            Terminate();
        }

        private void StartNew(string remoteSystem)
        {
            object[] arrParams = new[] { "notepad" };
            try
            {
                manageClass = new ManagementClass(myScope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
                manageClass.InvokeMethod("Create", arrParams);
                ConnectToRemoteMachine(remoteSystem);
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString())
                //Console.WriteLine(Information.Err.Description);
            }
        }

        private void Terminate()
        {
            try
            {
                string endProc = "notepad";
                foreach (ManagementObject obj in objSearcher.Get())
                {
                    string caption = obj.GetText(TextFormat.Mof).Trim();
                    if (caption.Contains(endProc.Trim()))
                        obj.InvokeMethod(opsObserver, "Terminate", null/* TODO Change to default(_) if this is not a reference type */);
                }
            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString())
                //Console.WriteLine(Information.Err.Description);
            }
        }

        public ActionResult RupDashboard()
        {
            DataTable dt = DBProc.GetResultasDataTbl("select JBM_AutoID,JBM_ID from JBM_Info i inner join JBM_CustomerMaster c on c.CustID=i.CustID where c.CustSN='RUP'", Session["sConnSiteDB"].ToString());
            ViewBag.JournalList = from a in dt.AsEnumerable() select new SelectListItem { Text = a["JBM_ID"].ToString(), Value = a["JBM_AutoID"].ToString() };
            return View();
        }

        public ActionResult getRupDashboardReport(string ReportType, string ArticleID, string JournalID)
        {
            try
            {
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ArticleID###", ArticleID).Replace("###JournalID###", JournalID);

                if (strQryText == "")
                {
                    return Json(new { result = "Error", data = "Report not configured. Please contact S/W Team" }, JsonRequestBehavior.AllowGet);
                }

                DataTable dt = new DataTable();

                dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = (dt.Rows.Count > 0) ? JsonConvert.SerializeObject(dt) : "";
                var jsonResult = Json(new { result = "Success", data = jsonString }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult getRupDashboardDetails(string ReportType, string ArticleID, string JournalID)
        {
            try
            {
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ArticleID###", ArticleID).Replace("###JournalID###", JournalID);

                if (ArticleID != "")
                {
                    strQryText = strQryText.Replace("###ArticleID###", ArticleID);
                }

                DataTable dt = new DataTable();

                dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = JsonConvert.SerializeObject(dt);
                return Json(jsonString, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json("", JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult getArticleList(string ArticleID, string JournalID, string ReportType)
        {
            try
            {
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry"))
                    .Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ArticleID###", ArticleID).Replace("###JournalID###", JournalID);
                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = from a in dt.AsEnumerable() select new[] { a[0].ToString() };
                return Json(jsonString, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("", JsonRequestBehavior.AllowGet);
            }

        }

        public ActionResult Event()
        {
            string strQryText = "SELECT DeptCode,DeptName FROM JBM_DepartmentMaster WHERE DeptSN in ('Software','CE','ML','Pag','PM','ML-FL','TechSupport','Vendor')";

            DataTable dTable = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

            ViewBag.ExternalEvent = getDashboard(onload: true);

            ViewBag.DeptList = from b in dTable.AsEnumerable() select new SelectListItem { Text = b["DeptName"].ToString(), Value = b["DeptCode"].ToString() };
            ViewBag.Title = Session["Event"].ToString();
            return View();
        }

        [SessionExpire]
        public dynamic getDashboard(string fromDate = "", string toDate = "", string custID = "", string[] jID = null, bool onload = false)
        {
            try
            {
                //string strQry = clsINI.GetProfileString("SQL_Query_Stage", "Event", Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ExternalEventID###", "").Replace("###Query###", "2").Replace("###Event###", Session["Event"].ToString()).Replace("###fromDate###", fromDate).Replace("###toDate###", toDate).Replace("###Status###", "0,1,2,4,5");
                string strCustIDQry = "", strJIDQry = "";
                
                if (custID !=null && custID != "")
                {
                    strCustIDQry = " AND i.CustID='" + custID + "'";
                }

                if (jID != null && jID.Length >= 0)
                {                                        
                    clsINI obj = new clsINI();
                    foreach (string strJID in jID)
                    {
                        if (strJIDQry != "")
                        {
                            strJIDQry = strJIDQry + ",";
                        }
                        strJIDQry = strJIDQry + obj.getJournlIDByTeam(Session["sCustAcc"].ToString(), strJID, Session["sConnSiteDB"].ToString());
                    }                  
                    strJIDQry = " AND i.JBM_AutoID in (" + strJIDQry + ")";
                }

                string strQry = clsINI.GetProfileString("SQL_Query_Stage", "Event", Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ExternalEventID###", "").Replace("###Query###", "2").Replace("###Event###", Session["Event"].ToString()).Replace("###fromDate###", fromDate).Replace("###toDate###", toDate).Replace("###Status###", " and (([Event ID]='EV003' and [Status] not in (99999)) or ([Event ID]<>'EV003' and [Status] in (0,1,2,4,5,400)))").Replace("###CustID###", strCustIDQry).Replace("###JID###", strJIDQry);

                DataTable dt = DBProc.GetResultasDataTbl(strQry, Session["sConnSiteDB"].ToString());

                var result = from a in dt.AsEnumerable() select new[] { a["Event Type"].ToString(), a["Event ID"].ToString(), a["Failed"].ToString(), a["Total"].ToString(), a["EventDate"].ToString(), a["Count_Failed"].ToString(), a["Count_Total"].ToString() };

                if (onload)
                {
                    return result;
                }

                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

        public ActionResult ExternalEvent()
        {
            Session["Event"] = "External";

            return RedirectToAction("Event");
        }

        public ActionResult getExternalEvent(string ReportType, string EventTypeID, string Status, string fromDate, string toDate, string custID = "", string[] jID = null)
        {
            try
            {
                string strCustIDQry = "", strJIDQry = "";

                if (custID != null && custID != "")
                {
                    strCustIDQry = " AND i.CustID='" + custID + "'";
                }

                if (jID != null && jID.Length >= 0)
                {
                    clsINI obj = new clsINI();
                    foreach (string strJID in jID)
                    {
                        if (strJIDQry != "")
                        {
                            strJIDQry = strJIDQry + ",";
                        }
                        strJIDQry = strJIDQry + obj.getJournlIDByTeam(Session["sCustAcc"].ToString(), strJID, Session["sConnSiteDB"].ToString());
                    }
                    strJIDQry = " AND i.JBM_AutoID in (" + strJIDQry + ")";
                }

                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ExternalEventID###", EventTypeID).Replace("###Query###", "1").Replace("###Event###", Session["Event"].ToString()).Replace("###fromDate###", fromDate).Replace("###toDate###", toDate).Replace("###Status###", Status).Replace("###CustID###", strCustIDQry).Replace("###JID###", strJIDQry);

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = (ds.Tables[0].Rows.Count > 0) ? JsonConvert.SerializeObject(ds.Tables[0]) : "";
                var eventSummary = (ds.Tables[1].Rows.Count > 0) ? JsonConvert.SerializeObject(ds.Tables[1]) : "";

                var jsonResult = Json(new { result = "Success", data = jsonString, summary = eventSummary }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;

            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult updateConversation(string autoID, string Message, List<HttpPostedFileBase> files)
        {
            try
            {
                if (Message != "")
                {
                    Message = Message + "<br/>";
                }

                if (files != null)
                {
                    Message = Message + EventFileUpload(autoID, Message, files);
                }

                string strQryText = "declare @xmlExists varchar(max) select @xmlExists=convert(varchar(max),[Conversation]) from " + Session["sCustAcc"].ToString() + "_ExternalEventAccess where AutoID=" + autoID + " if(Isnull(@xmlExists,'')='') begin UPDATE " + Session["sCustAcc"].ToString() + "_ExternalEventAccess SET [Conversation]='<conversation><message name=''" + Session["EmpName"].ToString() + "'' userid=''" + Session["UserID"].ToString() + "'' DeptName=''" + Session["DeptName"].ToString() + "'' >" + Message + "</message></conversation>' WHERE AutoID = " + autoID + " end else begin UPDATE " + Session["sCustAcc"].ToString() + "_ExternalEventAccess SET [Conversation].modify('insert <message name=''" + Session["EmpName"].ToString() + "'' userid=''" + Session["UserID"].ToString() + "'' DeptName=''" + Session["DeptName"].ToString() + "''>" + Message + "</message> into (/conversation)[1]') WHERE AutoID = " + autoID + " end select Isnull([Conversation],'') from " + Session["sCustAcc"].ToString() + "_ExternalEventAccess where AutoID=" + autoID + "";
                string result = DBProc.GetResultasString(strQryText, Session["sConnSiteDB"].ToString());
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

        private string EventFileUpload(string autoID, string Message, List<HttpPostedFileBase> files)
        {
            try
            {
                //string dir = "D:\\" + Session["Event"].ToString() + "\\Conversation\\" + autoID;
                string dir = Path.Combine(Session["InwardDirPath"].ToString(), Session["Event"].ToString(), "Conversation", autoID);
                string result = "";

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                foreach (HttpPostedFileBase file in files)
                {
                    string tempFileName = file.FileName;
                    string fileName = tempFileName;
                    int count = 0;
                    while (System.IO.File.Exists(Path.Combine(dir, fileName)))
                    {
                        count++;
                        fileName = Path.GetFileNameWithoutExtension(tempFileName) + " (" + count + ")" + Path.GetExtension(tempFileName);
                    }

                    file.SaveAs(Path.Combine(dir, fileName));
                    result = result + "<span>" + fileName + "  <i class=''fa fa-download'' data-filename=''" + fileName + "''></i></span>";
                }

                return result;
            }
            catch (Exception ex)
            {
                return "";
            }

        }

        public ActionResult EventFileDownload(string autoID, string fileName, bool status = false)
        {
            try
            {
                //string dir = "D:\\" + Session["Event"].ToString() + "\\Conversation\\" + autoID + "\\" + fileName;
                string dir = Path.Combine(Session["InwardDirPath"].ToString(), Session["Event"].ToString(), "Conversation", autoID);

                if (System.IO.File.Exists(dir))
                {

                    byte[] fileBytes = System.IO.File.ReadAllBytes(dir);

                    if (status)
                    {
                        return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
                    }
                    else
                    {
                        return Json("1");
                    }

                }
                else
                {
                    return Json("0");
                }

            }
            catch (Exception ex)
            {
                return Json("Error:" + ex.Message);
            }
        }

        public string DeptAllocation(string autoID, string deptCode)
        {
            try
            {
                string strQryText = "";

                if (deptCode == null)
                {
                    strQryText = "update " + Session["sCustAcc"].ToString() + "_ExternalEventAccess set EmpLogin='" + Session["UserID"].ToString() + "' where AutoID=" + autoID + "";
                }
                else
                {
                    strQryText = "update " + Session["sCustAcc"].ToString() + "_ExternalEventAccess set DeptCode='" + deptCode + "',EmpLogin=Null  where AutoID=" + autoID + "";
                }

                string result = DBProc.GetResultasString(strQryText, Session["sConnSiteDB"].ToString());

                if (result != "")
                {
                    return "success";
                }
                return "";
            }
            catch (Exception ex)
            {
                return "Error : " + ex.Message;
            }
        }

        private DataTable ExternalEventSummary(string EventTypeID)
        {
            try
            {
                string ReportType = "Event";
                string strQryText = clsINI.GetProfileString("SQL_Query_Stage", ReportType, Server.MapPath("~/Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ExternalEventID###", EventTypeID).Replace("###Query###", "3").Replace("###Event###", Session["Event"].ToString());

                DataTable dt = new DataTable();

                return DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());
            }
            catch (Exception ex)
            {
                return new DataTable();
            }
        }

        public string EventClose(string autoID)
        {
            try
            {
                string strQryText = "";
                if (Session["Event"].ToString() == "External")
                {
                    strQryText = "update " + Session["sCustAcc"].ToString() + "_ExternalEventAccess set Status=3  where AutoID=" + autoID;
                }
                else if (Session["Event"].ToString() == "Internal")
                {
                    strQryText = "update JBM_InternalAPIEvent set Status=3  where AutoID=" + autoID;
                }

                string result = DBProc.GetResultasString(strQryText, Session["sConnSiteDB"].ToString());
                return "success";
            }
            catch (Exception ex)
            {
                return "Error:" + ex.Message;
            }

        }

        public string EventRetrigger(string autoID)
        {
            try
            {//update TF_ExternalEventAccess set status =null where status=3 

                string strEventType = "";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select JBM_AutoID,AutoartID,ExternalEventID,ArticleID FROM " + Session["sCustAcc"].ToString() + "_ExternalEventAccess WHERE AutoID=" + autoID + "", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    strEventType = ds.Tables[0].Rows[0]["ExternalEventID"].ToString();
                }

                string strQryText = "";
                if (Session["Event"].ToString() == "External")
                {
                    strQryText = "UPDATE " + Session["sCustAcc"].ToString() + "_ExternalEventAccess set Status=0,MaxTry=0 where AutoID=" + autoID + " DECLARE @AutoArtID VARCHAR(50),@EventID VARCHAR(50) SELECT @AutoArtID=AutoArtID,@EventID=ExternalEventID FROM " + Session["sCustAcc"].ToString() + "_ExternalEventAccess WHERE AutoID=" + autoID + "";  // UPDATE " + Session["sCustAcc"].ToString() + "_ExternalEventAccess set Status=3 WHERE AutoArtID=@AutoArtID and ExternalEventID=@EventID and AutoID!=" + autoID + "
                }
                else if (Session["Event"].ToString() == "Internal")
                {
                    strQryText = "UPDATE JBM_InternalAPIEvent SET Status=0,MaxTry=0,EventStatus=NULL WHERE AutoID=" + autoID + " DECLARE @AutoArtID VARCHAR(50),@EventID VARCHAR(50) SELECT @AutoArtID=AutoArtID,@EventID=EventID FROM JBM_InternalAPIEvent WHERE AutoID=" + autoID + " UPDATE JBM_InternalAPIEvent set Status=3 WHERE AutoArtID=@AutoArtID and EventID=@EventID and AutoID!=" + autoID + "";
                }

                string result = DBProc.GetResultasString(strQryText, Session["sConnSiteDB"].ToString());

          
                string exeFilePath = "";

                try
                {

                    string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                    XmlNodeList objNodelist;
                    string strConnValue = string.Empty;
                    XmlDocument objxml = new XmlDocument();
                    objxml.Load(strPath);
                    if (objxml.InnerXml != "Nothing")
                    {
                        objNodelist = objxml.SelectNodes("//config/" + Session["sCustAcc"].ToString() + "/api-app-" + GlobalVariables.strEnvironment.ToString().ToLower() + "[@event-type='" + strEventType + "']");
                        if (objNodelist.Count > 0)
                        {
                            exeFilePath = objNodelist.Item(0).InnerText.ToString();
                        }
                    }
                }
                catch (Exception)
                {

                }

                String output = "success";
                //string exeFilePath = @"E:\KGL_Projects\MIS\Live\Web\trunk\AutoApplication\FP_DeprecationEvent\FP_DeprecationEvent\bin\Debug\FP_DeprecationEvent.exe";
                if (exeFilePath != "")
                {
                    Generic gen = new Generic();
                    gen.WriteLog("EXE Path: " + exeFilePath);
                    using (Process process = new Process())
                    {
                        process.StartInfo.FileName = exeFilePath;
                        process.StartInfo.Arguments = "AutoTrigger|" + autoID;
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.RedirectStandardOutput = true;
                        gen.WriteLog("Begin start");
                        process.Start();
                        gen.WriteLog("End start");
                        // Synchronously read the standard output of the spawned process.
                        StreamReader reader = process.StandardOutput;
                        output = reader.ReadToEnd();
                        gen.WriteLog("Result " + output);
                        // Write the redirected output to this application's window.

                        process.WaitForExit();
                    }

                    output = output.Replace("AutoTrigger", "");
                    output = Regex.Replace(output, @"\r\n?|\n", "");
                    gen.WriteLog("Output " + output);
                }

                return output.ToLower();
            }
            catch (Exception ex)
            {
                return "Error:" + ex.Message;
            }

        }

        public ActionResult InternalEvent()
        {
            Session["Event"] = "Internal";
            //string strQry = clsINI.GetProfileString("SQL_Query_Stage", "Event", Server.MapPath("../Query/SQLQueries.qry")).Replace("###CustAccess###", Session["sCustAcc"].ToString()).Replace("###ExternalEventID###", "").Replace("###Query###", "2").Replace("###Event###", Session["Event"].ToString());
            //string strQryText = "select DeptCode,DeptName from JBM_DepartmentMaster";

            //DataTable dt = DBProc.GetResultasDataTbl(strQry, Session["sConnSiteDB"].ToString());
            //DataTable dTable = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

            //ViewBag.ExternalEvent = from a in dt.AsEnumerable() select new[] { a["ExternalEventType"].ToString(), a["ExternalEventID"].ToString(), a["TotalRecord"].ToString(), a["AllEventTotal"].ToString(), a["Percentage"].ToString() };
            //ViewBag.DeptList = from b in dTable.AsEnumerable() select new SelectListItem { Text = b["DeptName"].ToString(), Value = b["DeptCode"].ToString() };

            return RedirectToAction("Event");
        }


        public ActionResult TransmittalEvent()
        {
            return View();
        }
        public ActionResult getTransmittalList()
        {
            try
            {
                try
                {
                    string strQueryFinal = "Select Stage as [Transmittal Type],CONVERT(varchar,ExternalEventActDate,9) as [Event Date],StatusDescription as [Status],ArticleID as [Article/Issue ID],FileName,Status as [Action],MaxTry from TF_ExternalEventAccess where externaleventid='EV003' and ExternalEventActDate >= DATEADD(day,-30, GETDATE()) order by ExternalEventActDate desc";

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {
                                         a[0].ToString(),
                                         a[1].ToString(),
                                         a[2].ToString(),
                                         a[3].ToString(),
                                         a[4].ToString(),
                                         CreateDynamicItem(a[5].ToString()!= "" ?Convert.ToInt32(a[5].ToString()):0,a[6].ToString()!= "" ?Convert.ToInt32(a[6].ToString()):0)
                                    };
                    return Json(new { dataResult = JSONString }, JsonRequestBehavior.AllowGet);
                }
                catch (Exception)
                {
                    return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public string CreateDynamicItem(int status, int maxtry)
        {
            string strResult = "";
            if (Convert.ToInt32(maxtry) > 0 && Convert.ToInt32(status) == 0)
            {
                strResult = "Fail";
            }
            return strResult.ToString();
        }
        public ActionResult DownloadTrigger(string ArticleUniqueID, string TransStage)
        {
            String output = "";
            string exeFilePath = @"C:\Applications\TME_Download_UAT\S3_TME_Download.exe";
            Generic gen = new Generic();
            gen.WriteLog("EXE Path: " + exeFilePath);
            using (Process process = new Process())
            {
                process.StartInfo.FileName = exeFilePath;
                ///process.StartInfo.Arguments = "AutoTrigger|" + autoID;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                gen.WriteLog("Begin start");
                process.Start();
                gen.WriteLog("End start");
                // Synchronously read the standard output of the spawned process.
                StreamReader reader = process.StandardOutput;
                output = reader.ReadToEnd();
                gen.WriteLog("Result " + output);
                // Write the redirected output to this application's window.

                process.WaitForExit();
            }

            //output = output.Replace("AutoTrigger", "");
            //output = Regex.Replace(output, @"\r\n?|\n", "");
            //gen.WriteLog("Output " + output);

            return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ArticleReport()
        {
            if (Session["sCustAcc"].ToString() == "BK" || Session["sCustAcc"].ToString() == "MG")
            {
                Init_Tables.gTblChapterOrArticleInfo = "_ChapterInfo";
            }
            else
            {
                Init_Tables.gTblChapterOrArticleInfo = "_ArticleInfo";
            }

            Session["strHomeURL"] = Request.Url.AbsoluteUri.ToString();
            //Session["sCustAcc"] = "KW";
            Session["employeeJournal"] = "";

            string strQuery = "if exists (SELECT CM.CustSN FROM JBM_EmployeeConfig EC INNER JOIN JBM_CustomerMaster CM ON CM.CustID=EC.Description WHERE EC.ProcessID='CID0001' AND EC.EmpAutoID='" + Session["EmpAutoId"].ToString() + "' AND CM.CustType='" + Session["sCustAcc"].ToString() + "' AND CM.Cust_Disabled IS NULL) begin SELECT CM.CustSN,CM.CustID FROM JBM_EmployeeConfig EC INNER JOIN JBM_CustomerMaster CM ON CM.CustID=EC.Description WHERE EC.ProcessID='CID0001' AND EC.EmpAutoID='" + Session["EmpAutoId"].ToString() + "' AND CM.CustType='" + Session["sCustAcc"].ToString() + "' AND CM.Cust_Disabled IS NULL ORDER BY CM.CustSN DESC end else begin select c.CustSN,c.CustID from JBM_CustomerMaster c where c.Cust_Disabled is NULL and c.CustType='" + Session["sCustAcc"].ToString() + "' order by c.CustSn desc end";
            DataTable dt = new DataTable();
            dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
            ViewBag.CustList = from a in dt.AsEnumerable() select new SelectListItem { Text = a["CustSN"].ToString(), Value = a["CustID"].ToString() };

            strQuery = "declare @employeeJournal varchar(max)='' If exists (Select JBM_Autoid from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.CustID=EC.Description and Ec.empautoid='" + Session["EmpAutoId"].ToString() + "' and ProcessID='CID0001' and EC.Description not in (Select CustID from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.JBM_AutoID=EC.Description and EC.empautoid='" + Session["EmpAutoId"].ToString() + "' and ProcessID='JID0001') Union Select JBM_Autoid from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.JBM_AutoID = EC.Description and EC.empautoid = '" + Session["EmpAutoId"].ToString() + "' and ProcessID = 'JID0001' and JI.JBM_Disabled=0) begin set @employeeJournal='Select JBM_Autoid from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.CustID=EC.Description and Ec.empautoid=''" + Session["EmpAutoId"].ToString() + "'' and ProcessID=''CID0001'' and EC.Description not in (Select CustID from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.JBM_AutoID=EC.Description and EC.empautoid=''" + Session["EmpAutoId"].ToString() + "'' and ProcessID=''JID0001'') Union Select JBM_Autoid from JBM_Info JI Inner Join JBM_EmployeeConfig EC on JI.JBM_AutoID = EC.Description and EC.empautoid = ''" + Session["EmpAutoId"].ToString() + "'' and ProcessID = ''JID0001'' and JI.JBM_Disabled=0' end select @employeeJournal";

            dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

            if (dt.Rows.Count > 0)
            {
                Session["employeeJournal"] = dt.Rows[0][0].ToString();
            }

            return View();
        }

        public dynamic getJournalList(string custID = "", bool convertToJSON = true)
        {
            try
            {
                string strCustID = "";

                if (custID != "")
                {
                    strCustID = " AND c.CustID='" + custID + "'";
                }

                string strQuery = "if exists (select i.JBM_AutoID from JBM_EmployeeConfig cc inner join JBM_Info i on i.JBM_AutoID=cc.Description inner join JBM_CustomerMaster c on c.CustID=i.CustID where cc.ProcessID='JID0001' and EmpAutoID='" + Session["EmpAutoId"].ToString() + "' and c.CustType='" + Session["sCustAcc"].ToString() + "' and c.Cust_Disabled is null and JBM_Disabled='0' " + strCustID + " ) begin select i.JBM_AutoID,i.JBM_ID  from JBM_EmployeeConfig cc inner join JBM_Info i on i.JBM_AutoID=cc.Description inner join JBM_CustomerMaster c on c.CustID=i.CustID where cc.ProcessID='JID0001' and EmpAutoID='" + Session["EmpAutoId"].ToString() + "' and c.CustType='" + Session["sCustAcc"].ToString() + "' and c.Cust_Disabled is null and JBM_Disabled='0' " + strCustID + " order by i.JBM_ID end else begin select i.JBM_AutoID,i.JBM_ID  from JBM_Info i inner join JBM_CustomerMaster c on c.CustID=i.CustID where c.CustType='" + Session["sCustAcc"].ToString() + "' and c.Cust_Disabled is null and JBM_Disabled='0'" + strCustID + " order by i.JBM_ID end";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                var result = from a in dt.AsEnumerable() select new SelectListItem { Text = a["JBM_ID"].ToString(), Value = a["JBM_AutoID"].ToString() };
                if (convertToJSON)
                {
                    return Json(result);
                }
                else
                {
                    return result;
                }
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public ActionResult getArticle_AutoComplete(string ArticleID, string JournalID, string custID)
        {
            try
            {
                //string strQryText = "Select ChapterID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " where ChapterID like '" + ArticleID + "%'";
                string strJournal = "";
                string strCustID = "";

                if (JournalID != "")
                {
                    strJournal = " and i.JBM_AutoID='" + JournalID + "' ";
                }
                else if (Session["employeeJournal"].ToString() != "")
                {
                    strJournal = " and i.JBM_AutoID in (" + Session["employeeJournal"].ToString() + ") ";
                }

                if (custID != "")
                {
                    strCustID = " and i.CustID='" + custID + "' ";
                }

                //string strQryText = "declare @ArticleInfo table (ArticleID varchar(100)) if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select ChapterID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where ChapterID like '" + ArticleID + "%' " + strJournal + strCustID + " if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select DOI from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + "  a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where DOI like '" + ArticleID + "%' " + strJournal + strCustID + " if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select AutoArtID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where AutoArtID like '" + ArticleID + "%' " + strJournal + strCustID + " if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select IntrnlID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where IntrnlID like '" + ArticleID + "%' " + strJournal + strCustID + " end end end end select distinct * from @ArticleInfo";

                string strQryText = "declare @ArticleInfo table (ArticleID varchar(100)) if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select ChapterID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where ChapterID like '" + ArticleID + "%' " + strJournal + strCustID + " if not exists(select * from @ArticleInfo) begin insert into @ArticleInfo select DOI from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + "  a inner join JBM_Info i on i.JBM_AutoID=a.JBM_AutoID where DOI like '" + ArticleID + "%' " + strJournal + strCustID + " end end select distinct * from @ArticleInfo";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                var jsonString = from a in dt.AsEnumerable() select new[] { a[0].ToString() };
                return Json(jsonString, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("", JsonRequestBehavior.AllowGet);
            }

        }

        public ActionResult downloadArticleFiles(string path, string fileName, bool status = false)
        {
            try
            {

                if (System.IO.File.Exists(path))
                {

                    if (status)
                    {
                        byte[] fileBytes = System.IO.File.ReadAllBytes(path);
                        return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
                    }
                    else
                    {
                        return Json("1");
                    }

                }
                else
                {
                    return Json("0");
                }

            }
            catch (Exception ex)
            {
                return Json("Error:" + ex.Message);
            }
        }
        public ActionResult ArticleReport_getArticleDetails(string ArticleID, string JournalID, string custID)
        {
            try
            {
                string strJournal = "";
                string strCustID = "";

                if (JournalID != "")
                {
                    strJournal = " and i.JBM_AutoID='" + JournalID + "'";
                }
                else if (Session["employeeJournal"].ToString() != "")
                {
                    strJournal = " and i.JBM_AutoID in (" + Session["employeeJournal"].ToString() + ") ";
                }

                if (custID != "")
                {
                    strCustID = " and c.CustID='" + custID + "' ";
                }

                string strQuery = "select top 1 a.AutoID,a.JBM_AutoID,i.JBM_Intrnl,a.IntrnlId,c.CustSN,i.JBM_ID as [Journal Name],i.SiteID,a.AutoArtID,ChapterID as [Article ID],p.ArticleTitle as [Article Title],t.ArtTypeDesc as [Article Type],iss as [Vol:Iss],corraut as [Corresponding Author],convert(varchar(11),s.CeDispDate,106) as [CEDate],convert(varchar(11),s.Dispatchdate,106) as [ProofDate],convert(varchar(11),s.CrsTypeAutRec,106) as [ACDate] from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join " + Init_Tables.gTblJrnlInfo + " i on i.JBM_AutoID=a.JBM_AutoID inner join " + Init_Tables.gTblArticleTypes + " t on t.ArtTypeID=a.ArtTypeID inner join " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " s on s.AutoArtID=a.AutoArtID inner join " + Session["sCustAcc"].ToString() + "_ProdInfo" + " p on p.AutoArtID=a.AutoArtID inner join JBM_CustomerMaster c on c.CustID=i.CustID where (a.AutoArtID='" + ArticleID + "' or a.ChapterID='" + ArticleID + "' or a.IntrnlID='" + ArticleID + "' or a.DOI='" + ArticleID + "') and s.RevFinStage='FP' " + strJournal + strCustID + " order by  case when a.ChapterID='" + ArticleID + "' then 4 when a.DOI='" + ArticleID + "' then 3 when a.AutoArtID='" + ArticleID + "' then 2 when a.IntrnlID='" + ArticleID + "' then 1 else 0 end desc";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                List<string[]> fileArray = new List<string[]> { };

                if (dt.Rows.Count > 0)
                {
                    fileArray = ArticleReport_getFiles(dt);
                }

                var jsonData = (dt.Rows.Count > 0) ? JsonConvert.SerializeObject(dt) : "";

                var jsonResult = Json(new { result = "Success", data = jsonData, file = fileArray }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;
                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public List<string[]> ArticleReport_getFiles(DataTable dt)
        {
            try
            {

                RL.ArtDet art = new RL.ArtDet();
                string strIssue = dt.Rows[0]["Vol:Iss"].ToString();

                art.AD.ArticleID = dt.Rows[0]["Article ID"].ToString();
                art.AD.JAutoID = dt.Rows[0]["JBM_AutoID"].ToString();
                art.AD.InternalJID = dt.Rows[0]["JBM_Intrnl"].ToString();
                art.AD.InternalID = dt.Rows[0]["IntrnlId"].ToString();
                art.AD.AutoArtID = dt.Rows[0]["AutoArtID"].ToString();
                art.AD.CustSN = dt.Rows[0]["CustSN"].ToString();
                art.AD.CustAccess = Session["sCustAcc"].ToString();
                //art.AD.CustGroup = dt.Rows[0]["CustGroup"].ToString();
                art.AD.JrnlSiteID = dt.Rows[0]["SiteID"].ToString();
                art.AD.JID = dt.Rows[0]["Journal Name"].ToString();
                art.AD.ConnString = DBProc.getConnection(Session["sConnSiteDB"].ToString()).ConnectionString;

                string strInternalJID = art.AD.InternalJID.Contains("#") ? art.AD.JID : art.AD.InternalJID;

                string strVolNo = "000";
                string strIssNo = "00";

                if (strIssue != "")
                {
                    Proc_Split_VolIss(strIssue, ref strVolNo, ref strIssNo);
                }

                art.AD.VolDir = "Vol" + strVolNo + strIssNo;
                Generic gen = new Generic();

                List<string[]> strFileList = new List<string[]>();
                string[] strFile = new string[8];
                string strRoot = Proc_Get_Directory_Path(ref art, "F20-01", false);
                gen.WriteLog("Path... " + strRoot);
                string strFilePath = "";
                string strInputFileName = "";
                bool hasFile = false;

                string CEDate = dt.Rows[0]["CEDate"].ToString();
                string ProofDate = dt.Rows[0]["ProofDate"].ToString();
                string ACDate = dt.Rows[0]["ACDate"].ToString();

                try
                {
                    strFilePath = Path.Combine(strRoot, "doc");
                    if (Directory.Exists(strFilePath) & !Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(pnas|asn|aapg|aai|ada)"))
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                        FileInfo[] fileInfo = dirInfo.GetFiles("*.doc*", SearchOption.TopDirectoryOnly);

                        foreach (FileInfo fi in fileInfo)
                        {
                            if (fi.Extension == ".doc" | fi.Extension == ".docx")
                            {
                                strFile = new string[8];
                                strFile[0] = fi.Name;
                                strFile[1] = "Original Manuscript";
                                strFile[2] = fi.LastWriteTime.ToString("dd MMM yyyy");
                                strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                strFile[4] = "";
                                strFile[5] = fi.Name;
                                strFile[6] = fi.Name;
                                strFile[7] = fi.FullName;
                                strFileList.Add(strFile);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("Log1... " + ex.Message);
                }

                try
                {
                    strFilePath = Path.Combine(strRoot);
                    if (Directory.Exists(strFilePath) & Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(ada)"))
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                        //FileInfo[] fileInfo = dirInfo.GetFiles("*metadata.html", SearchOption.TopDirectoryOnly);
                        FileInfo[] fileInfo = dirInfo.GetFiles(art.AD.JID + "_" + art.AD.ArticleID + ".html", SearchOption.TopDirectoryOnly);

                        foreach (FileInfo fi in fileInfo)
                        {
                            strFile = new string[8];
                            strFile[0] = fi.Name;
                            strFile[1] = "Cover Sheet";
                            strFile[2] = fi.LastWriteTime.ToString("dd MMM yyyy");
                            strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                            strFile[4] = "";
                            strFile[5] = fi.Name;
                            strFile[6] = fi.Name;
                            strFile[7] = fi.FullName;
                            strFileList.Add(strFile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("Log1... " + ex.Message);
                }

                try
                {
                    strFilePath = Path.Combine(strRoot, "pdf");
                    if (Directory.Exists(strFilePath) & Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(ada)"))
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                        FileInfo[] fileInfo = dirInfo.GetFiles("*.pdf", SearchOption.TopDirectoryOnly);

                        foreach (FileInfo fi in fileInfo)
                        {
                            strFile = new string[8];
                            strFile[0] = fi.Name;
                            strFile[1] = "Original Manuscript PDF";
                            strFile[2] = fi.LastWriteTime.ToString("dd MMM yyyy");
                            strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                            strFile[4] = "";
                            strFile[5] = fi.Name;
                            strFile[6] = fi.Name;
                            strFile[7] = fi.FullName;
                            strFileList.Add(strFile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("Log1... " + ex.Message);
                }

                try
                {
                    if (!Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(asn|aai)"))
                    {
                        strRoot = Proc_Get_Directory_Path(ref art, "F40-01", false);
                        strFilePath = Path.Combine(strRoot, "Incoming");
                        string CEFileName = "FIN";

                        if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
                        {
                            strInputFileName = art.AD.CustAccess + "-" + art.AD.CustSN + "-" + art.AD.InternalJID + art.AD.InternalID;
                        }
                        else
                        {
                            strInputFileName = art.AD.CustAccess + "-" + art.AD.InternalJID + art.AD.InternalID;
                        }

                        if (Session["UserID"].ToString().ToLower() == "ada")
                        {
                            CEFileName = "CE";
                            strInputFileName = strInputFileName + "_" + CEFileName;
                        }
                        else if (Session["UserID"].ToString().ToLower() == "aapg")
                        {
                            CEFileName = "PER";
                        }

                        if (Directory.Exists(strFilePath))
                        {
                            DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                            FileInfo[] fileInfo = dirInfo.GetFiles(strInputFileName + ".doc*", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fi in fileInfo)
                            {
                                if (fi.Extension == ".doc" | fi.Extension == ".docx")
                                {
                                    string ext = fi.Extension.Contains(".docx") ? ".docx" : ".doc";
                                    strFile = new string[8];
                                    strFile[0] = art.AD.JID + art.AD.ArticleID + "_" + CEFileName + ext;
                                    strFile[1] = "" + CEFileName + " Manuscript";
                                    strFile[2] = CEDate;//fi.LastWriteTime.ToString("dd MMM yyyy");
                                    strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                    strFile[4] = "";
                                    strFile[5] = fi.Name;
                                    strFile[6] = art.AD.JID + art.AD.ArticleID + "_" + CEFileName + ext;
                                    strFile[7] = fi.FullName;
                                    strFileList.Add(strFile);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    gen.WriteLog("Log2... " + ex.Message);
                }

                try
                {
                    strRoot = Proc_Get_Directory_Path(ref art, "F250-01", false);
                    strFilePath = strRoot.Replace("IProof", "sProof");
                    strFilePath = Path.Combine(strFilePath, "AEXP");
                    hasFile = false;

                    strInputFileName = art.AD.JID + art.AD.ArticleID + "_proof.pdf";

                    if (!Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(asn)"))
                    {
                        if (Directory.Exists(strFilePath))
                        {
                            DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                            FileInfo[] fileInfo = dirInfo.GetFiles("*.zip", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fi in fileInfo)
                            {
                                if (fi.Extension == ".zip")
                                {
                                    Ionic.Zip.ZipFile oZip = new Ionic.Zip.ZipFile();
                                    oZip = Ionic.Zip.ZipFile.Read(fi.FullName);
                                    string extractFolder = Path.Combine(fi.DirectoryName, fi.Name.Substring(0, fi.Name.Length - 4));

                                    if (!Directory.Exists(extractFolder))
                                    {
                                        Directory.CreateDirectory(extractFolder);
                                    }
                                    else
                                    {
                                        List<FileInfo> fList = new DirectoryInfo(extractFolder).GetFiles("*.*", SearchOption.AllDirectories).ToList();
                                        foreach (FileInfo fInfo in fList)
                                        {
                                            fInfo.Delete();
                                        }
                                    }

                                    oZip.ExtractAll(extractFolder, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                                    oZip.Dispose();

                                    DirectoryInfo zipDirInfo = new DirectoryInfo(extractFolder);
                                    FileInfo[] zipFileInfo = zipDirInfo.GetFiles(strInputFileName, SearchOption.AllDirectories);

                                    foreach (FileInfo zipFi in zipFileInfo)
                                    {
                                        strFile = new string[8];
                                        strFile[0] = art.AD.JID + art.AD.ArticleID + "_eproof" + zipFi.Extension;
                                        strFile[1] = "Author Proof";
                                        strFile[2] = ProofDate;//zipFi.LastWriteTime.ToString("dd MMM yyyy");
                                        strFile[3] = Convert.ToInt32((zipFi.Length) / 1024) + " KB";
                                        strFile[4] = "";
                                        strFile[5] = zipFi.Name;
                                        strFile[6] = art.AD.JID + art.AD.ArticleID + "_eproof" + zipFi.Extension;
                                        strFile[7] = zipFi.FullName;
                                        hasFile = true;
                                        strFileList.Add(strFile);
                                    }
                                }
                            }
                        }

                        strFilePath = strRoot;
                        if (Directory.Exists(strFilePath) & hasFile == false)
                        {
                            DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                            FileInfo[] fileInfo = dirInfo.GetFiles("*.zip", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fi in fileInfo)
                            {
                                if (fi.Extension == ".zip")
                                {
                                    Ionic.Zip.ZipFile oZip = new Ionic.Zip.ZipFile();
                                    oZip = Ionic.Zip.ZipFile.Read(fi.FullName);
                                    string extractFolder = Path.Combine(fi.DirectoryName, fi.Name.Substring(0, fi.Name.Length - 4));

                                    if (!Directory.Exists(extractFolder))
                                    {
                                        Directory.CreateDirectory(extractFolder);
                                    }
                                    else
                                    {
                                        List<FileInfo> fList = new DirectoryInfo(extractFolder).GetFiles("*.*", SearchOption.AllDirectories).ToList();
                                        foreach (FileInfo fInfo in fList)
                                        {
                                            fInfo.Delete();
                                        }
                                    }

                                    oZip.ExtractAll(extractFolder, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                                    oZip.Dispose();

                                    DirectoryInfo zipDirInfo = new DirectoryInfo(extractFolder);
                                    FileInfo[] zipFileInfo = zipDirInfo.GetFiles(strInputFileName, SearchOption.AllDirectories);

                                    foreach (FileInfo zipFi in zipFileInfo)
                                    {
                                        strFile = new string[8];
                                        strFile[0] = art.AD.JID + art.AD.ArticleID + "_eproof" + zipFi.Extension;
                                        strFile[1] = "Author Proof";
                                        strFile[2] = ProofDate;//zipFi.LastWriteTime.ToString("dd MMM yyyy");
                                        strFile[3] = Convert.ToInt32((zipFi.Length) / 1024) + " KB";
                                        strFile[4] = "";
                                        strFile[5] = zipFi.Name;
                                        strFile[6] = art.AD.JID + art.AD.ArticleID + "_eproof" + zipFi.Extension;
                                        strFile[7] = zipFi.FullName;
                                        strFileList.Add(strFile);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                try
                {
                    strRoot = Proc_Get_Directory_Path(ref art, "F20-02", false);
                    strFilePath = Path.Combine(strRoot, "Rev1", "AuthorCorr");
                    string addFolder = Path.Combine(strRoot, "Rev1", strInternalJID + art.AD.ArticleID + "_Aucx");

                    if (Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(asn|aapg)"))
                    {
                        if (Directory.Exists(strFilePath))
                        {
                            List<FileInfo> addFileInfo = new DirectoryInfo(strFilePath).EnumerateFiles().Where(i => !i.Name.EndsWith(".xml") & !i.Name.EndsWith(".txt")).ToList();
                            string zipFilePath = "";

                            if (addFileInfo.Count > 1)
                            {

                                if (!Directory.Exists(addFolder))
                                {
                                    Directory.CreateDirectory(addFolder);
                                }
                                else
                                {
                                    List<FileInfo> fList = new DirectoryInfo(addFolder).GetFiles("*.*", SearchOption.AllDirectories).ToList();
                                    foreach (FileInfo fInfo in fList)
                                    {
                                        fInfo.Delete();
                                    }
                                }

                                foreach (FileInfo fileInfo in addFileInfo)
                                {
                                    string folderName = fileInfo.DirectoryName.ToString().Replace(strFilePath, addFolder);

                                    if (!Directory.Exists(folderName))
                                    {
                                        Directory.CreateDirectory(folderName);
                                    }

                                    if (fileInfo.Extension != ".txt" & fileInfo.Extension != ".xml")
                                    {
                                        fileInfo.CopyTo(Path.Combine(folderName, fileInfo.Name));
                                    }
                                }
                                zipFilePath = addFolder + ".zip";
                                Ionic.Zip.ZipFile outZipFiles = new Ionic.Zip.ZipFile();
                                outZipFiles.AddDirectory(addFolder);
                                outZipFiles.Save(zipFilePath);
                            }
                            else if (addFileInfo.Count == 1)
                            {
                                zipFilePath = addFileInfo[0].FullName;
                            }

                            if (System.IO.File.Exists(zipFilePath))
                            {
                                FileInfo fi = new FileInfo(zipFilePath);
                                strFile = new string[8];
                                strFile[0] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                strFile[1] = "Author Corrections 1";
                                strFile[2] = ACDate;//fi.LastWriteTime.ToString("dd MMM yyyy");
                                strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                strFile[4] = "";
                                strFile[5] = fi.Name;
                                strFile[6] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                strFile[7] = fi.FullName;
                                strFileList.Add(strFile);
                            }
                        }
                    }
                    else if (!Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(aai)"))
                    {

                        strInputFileName = art.AD.ArticleID + "_au.pdf";
                        hasFile = false;

                        if (Directory.Exists(strFilePath))
                        {
                            DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                            FileInfo[] fileInfo = dirInfo.GetFiles("*.pdf", SearchOption.TopDirectoryOnly);

                            foreach (FileInfo fi in fileInfo)
                            {
                                if (fi.Name.ToLower() == strInputFileName.ToLower() & fi.Extension == ".pdf")
                                {
                                    strFile = new string[8];
                                    strFile[0] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                    strFile[1] = "Author Corrections 1";
                                    strFile[2] = ACDate;//fi.LastWriteTime.ToString("dd MMM yyyy");
                                    strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                    strFile[4] = "";
                                    strFile[5] = fi.Name;
                                    strFile[6] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                    strFile[7] = fi.FullName;
                                    hasFile = true;
                                    strFileList.Add(strFile);
                                }
                            }
                        }

                        strFilePath = Path.Combine(strRoot, "Rev1");

                        if (Directory.Exists(strFilePath) & hasFile == false)
                        {
                            DirectoryInfo dirInfo = new DirectoryInfo(strFilePath);
                            FileInfo[] fileInfo = dirInfo.GetFiles("*.pdf", SearchOption.TopDirectoryOnly);

                            if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
                            {
                                strInputFileName = art.AD.CustAccess + "-" + art.AD.CustSN + "-" + art.AD.InternalJID + art.AD.InternalID + "_cx.pdf";
                            }
                            else
                            {
                                strInputFileName = art.AD.CustAccess + "-" + art.AD.InternalJID + art.AD.InternalID + "_cx.pdf";
                            }

                            foreach (FileInfo fi in fileInfo)
                            {
                                if (fi.Name.ToLower() == strInputFileName.ToLower() & fi.Extension == ".pdf")
                                {
                                    strFile = new string[8];
                                    strFile[0] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                    strFile[1] = "Author Corrections 1";
                                    strFile[2] = ACDate;//fi.LastWriteTime.ToString("dd MMM yyyy");
                                    strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                    strFile[4] = "";
                                    strFile[5] = fi.Name;
                                    strFile[6] = art.AD.JID + art.AD.ArticleID + "_Aucx" + fi.Extension;
                                    strFile[7] = fi.FullName;
                                    strFileList.Add(strFile);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                try
                {
                    if (!Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(pnas|aai)"))
                    {
                        strInputFileName = art.AD.ArticleID + "_pe.pdf";
                        strFilePath = strRoot;

                        string strRoot2 = Proc_Get_Directory_Path(ref art, "F250-01", true);
                        string strFilePath2 = "";

                        //if (Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(ada|aapg)"))
                        //{
                        //    strFilePath2 = strRoot2.Replace("IProof", "Rev");
                        //}
                        //else
                        //{
                        //    strFilePath2 = strRoot2.Replace("IProof", "sProof");
                        //    strFilePath2 = Path.Combine(strFilePath2, "AEXPRQC");
                        //}

                        string strInputFileName2 = "";
                        string strInputFileName3 = "";

                        DirectoryInfo dirInfo = new DirectoryInfo(strRoot);
                        List<DirectoryInfo> dirInfoList = dirInfo.EnumerateDirectories().OrderBy(d => d.Name).Where(i => i.Name.StartsWith("Rev")).ToList();

                        foreach (DirectoryInfo diInfo in dirInfoList)
                        {
                            string count = diInfo.Name.ToString().ToLower().Replace("rev", "");

                            try
                            {
                                int a = Convert.ToInt16(count);
                            }
                            catch (Exception ex)
                            {
                                count = "";
                            }

                            if (count != "")
                            {
                                strFilePath = Path.Combine(diInfo.FullName, "PECorr", strInputFileName);

                                if (System.IO.File.Exists(strFilePath))
                                {
                                    FileInfo fi = new FileInfo(strFilePath);
                                    strFile = new string[8];
                                    strFile[0] = art.AD.JID + art.AD.ArticleID + "_pe" + count + fi.Extension;
                                    strFile[1] = "Proof Correction " + count;
                                    strFile[2] = fi.LastWriteTime.ToString("dd MMM yyyy");
                                    strFile[3] = Convert.ToInt32((fi.Length) / 1024) + " KB";
                                    strFile[4] = "";
                                    strFile[5] = fi.Name;
                                    strFile[6] = art.AD.JID + art.AD.ArticleID + "_pe" + count + fi.Extension;//fi.Name;
                                    strFile[7] = fi.FullName;
                                    strFileList.Add(strFile);
                                }
                                else
                                {
                                    continue;
                                }

                                string oZipFileName = "";
                                string extractFolder = "";
                                hasFile = false;
                                if (Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(ada|aapg)"))
                                {
                                    strInputFileName2 = art.AD.JID + "_ART_CX_" + art.AD.ArticleID + "_" + count + "REV";
                                    strInputFileName3 = art.AD.JID + art.AD.ArticleID + "_rp" + count + ".pdf";
                                                                        
                                    strFilePath2 = Path.Combine(Path.GetDirectoryName(strRoot2), "Rev");
                                    oZipFileName = Path.Combine(strFilePath2, "Rev" + count, strInputFileName2 + ".zip");
                                    extractFolder = Path.Combine(strFilePath2, "Rev" + count, strInputFileName2);

                                    if (System.IO.File.Exists(oZipFileName))
                                    {
                                        hasFile = true;
                                    }                             
                                }

                                if (!Regex.IsMatch(Session["UserID"].ToString().ToLower(), "^(aapg)$") && !hasFile)
                                {
                                    strInputFileName2 = art.AD.JID + "_ART_AEXPRQC_" + art.AD.ArticleID + "_" + count + "REV";
                                    strInputFileName3 = art.AD.JID + art.AD.ArticleID + "_rp" + count + ".pdf";

                                    strFilePath2 = Path.Combine(Path.GetDirectoryName(strRoot2), "sProof", "AEXPRQC");
                                    oZipFileName = Path.Combine(strFilePath2, strInputFileName2 + ".zip");
                                    extractFolder = Path.Combine(strFilePath2, strInputFileName2);
                                }

                                //if (Regex.IsMatch(Session["UserID"].ToString().ToLower(), "(ada|aapg)"))
                                //{
                                //    strInputFileName2 = art.AD.JID + "_ART_CX_" + art.AD.ArticleID + "_" + count + "REV";
                                //    strInputFileName3 = art.AD.JID + art.AD.ArticleID + "_rp" + count + ".pdf";

                                //    oZipFileName = Path.Combine(strFilePath2, "Rev" + count, strInputFileName2 + ".zip");
                                //    extractFolder = Path.Combine(strFilePath2, "Rev" + count, strInputFileName2);
                                //}
                                //else
                                //{
                                //    strInputFileName2 = art.AD.JID + "_ART_AEXPRQC_" + art.AD.ArticleID + "_" + count + "REV";
                                //    strInputFileName3 = art.AD.JID + art.AD.ArticleID + "_rp" + count + ".pdf";

                                //    oZipFileName = Path.Combine(strFilePath2, strInputFileName2 + ".zip");
                                //    extractFolder = Path.Combine(strFilePath2, strInputFileName2);
                                //}


                                if (System.IO.File.Exists(oZipFileName))
                                {
                                    Ionic.Zip.ZipFile oZip = new Ionic.Zip.ZipFile();
                                    oZip = Ionic.Zip.ZipFile.Read(oZipFileName);

                                    if (!Directory.Exists(extractFolder))
                                    {
                                        Directory.CreateDirectory(extractFolder);
                                    }
                                    else
                                    {
                                        List<FileInfo> fList = new DirectoryInfo(extractFolder).GetFiles("*.*", SearchOption.AllDirectories).ToList();
                                        foreach (FileInfo fInfo in fList)
                                        {
                                            fInfo.Delete();
                                        }
                                    }

                                    oZip.ExtractAll(extractFolder, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                                    oZip.Dispose();

                                    DirectoryInfo zipDirInfo = new DirectoryInfo(extractFolder);
                                    FileInfo[] zipFileInfo = zipDirInfo.GetFiles(strInputFileName3, SearchOption.AllDirectories);

                                    foreach (FileInfo zipFi in zipFileInfo)
                                    {
                                        strFile = new string[8];
                                        strFile[0] = zipFi.Name;
                                        strFile[1] = "Revised Proof " + count;
                                        strFile[2] = zipFi.LastWriteTime.ToString("dd MMM yyyy");
                                        strFile[3] = Convert.ToInt32((zipFi.Length) / 1024) + " KB";
                                        strFile[4] = "";
                                        strFile[5] = zipFi.Name;
                                        strFile[6] = zipFi.Name;
                                        strFile[7] = zipFi.FullName;
                                        strFileList.Add(strFile);
                                    }
                                }
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                }
                
                return strFileList;
            }
            catch (Exception ex)
            {
                return new List<string[]>();
            }
        }

        private string Proc_Get_Directory_Path(ref RL.ArtDet A, string FolderIndex, Boolean blnIgnoreSlash)
        {
            string strRootPath = "";
            string strFolderDir = "";
            string strPath = "";

            try
            {
                DataRow[] dr = DBProc.GetResultasDataTbl("Select b.RootPath, a.FolderDir, a.CustType, a.FolderIndex , (Select K.Rootpath from JBM_RootDirectory K where K.RootID=a.ProPdfDir) as [ProPdfDir] from JBM_DeptFolders a, JBM_RootDirectory b where a.RootID=b.RootID and CustType='" + A.AD.CustAccess + "' and FolderIndex='" + FolderIndex + "'", Session["sConnSiteDB"].ToString()).Select();

                if (dr.Length != 0)
                {
                    string strLocation = A.AD.JrnlSiteID;
                    DataRow[] drJrnlSite = null;
                    string NewFolderStruct = string.Empty;

                    //if (ArtPickUpSite != "")
                    //    strLocation = Strings.Right(ArtPickUpSite, 5); // L0001-L0004
                    //else if (JBM_SiteID != "")
                    //    strLocation = JBM_SiteID;
                    //else if (JBM_SiteID == "")
                    //{
                    //    drJrnlSite = DBProc.GetResultasDataTbl("Select i.ColorID, c.CustSN, i.JBM_ID, i.JBM_AutoID, c.CustType, i.SiteID, i.JBM_Intrnl, i.JBM_NewFolder  from JBM_Info i, JBM_CustomerMaster c where i.JBM_Disabled <> '1' and c.CustID=i.CustID and CustType='" + strCustAccess + "' and JBM_Intrnl='" + strIntrnlJID + "'", "TempTable").Select();



                    //    strLocation = drJrnlSite[5].ToString();
                    //    NewFolderStruct = drJrnlSite[7].ToString();
                    //}


                    if (strLocation == "L0001" & dr[0][4].ToString() != "")
                        strRootPath = dr[0][4].ToString();
                    else
                        strRootPath = dr[0][0].ToString();


                    DataTable dt = DBProc.GetResultasDataTbl("Select SiteID, RootPath from JBM_SiteMaster where SiteID='" + strLocation + "'", Session["sConnSiteDB"].ToString());


                    foreach (DataRow drRow in dt.Rows)
                        strLocation = @"\\" + drRow["RootPath"].ToString() + @"\";


                    if (strLocation != "" & strRootPath.IndexOf(strLocation, StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        strRootPath = Strings.Replace(strRootPath, @"\\blrnas3\", strLocation, 1, -1, CompareMethod.Text);
                        strRootPath = Strings.Replace(strRootPath, @"\\chenas03\", strLocation, 1, -1, CompareMethod.Text);
                    }

                    strFolderDir = dr[0][1].ToString();

                    strFolderDir = Strings.Replace(strFolderDir, "###JID###", A.AD.InternalJID, 1, -1, CompareMethod.Text);
                    strFolderDir = Strings.Replace(strFolderDir, "###VOLDIR###", A.AD.VolDir, 1, -1, CompareMethod.Text);
                    strFolderDir = Strings.Replace(strFolderDir, "###INTRNLID###", A.AD.InternalID, 1, -1, CompareMethod.Text);
                    strFolderDir = Strings.Replace(strFolderDir, "###CUSTID###", A.AD.CustSN, 1, -1, CompareMethod.Text);
                    strFolderDir = Strings.Replace(strFolderDir, "###DOI###", A.AD.DOI, 1, -1, CompareMethod.Text);
                    strFolderDir = Strings.Replace(strFolderDir, "###ArticleID###", A.AD.ArticleID, 1, -1, CompareMethod.Text);
                    if (blnIgnoreSlash & Strings.Right(strFolderDir, 1) == @"\")
                        strFolderDir = Strings.Mid(strFolderDir, 1, Strings.Len(strFolderDir) - 1);

                    strPath = strRootPath + strFolderDir;
                    strPath = Strings.Replace(strPath, @"\&nbsp;\", @"\", 1, -1, CompareMethod.Text);
                    strPath = Strings.Replace(strPath, @"\ \", @"\", 1, -1, CompareMethod.Text);

                    //if (SessionHandler.gUserSiteId != Init_SiteID.gSiteBangalore & Strings.InStr(1, strPath, @"\\blrnas3\", Constants.vbTextCompare) > 0)
                    //{
                    //}
                    //else if ((strCustDir == "EMP" | strCustDir == "TNF" | strCustDir == "OUP") & (strIntrnlJID == "REIO-180374" | strIntrnlJID == "CHARLES_33564-180425" | strIntrnlJID == "ARMSTRONG-130610") & strCustAccess == "BK" & (HttpContext.Current.Request.Url.ToString().ToLower().Contains("10.18") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("localhost") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("cps.kwglobal.com") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("smarttrack.cenveo.com/smarttrack-ch") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("smarttrack.cenveo.com/smarttrack_test") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("smarttrack.kwglobal.com/smarttrack-ch") | HttpContext.Current.Request.Url.ToString().ToLower().Contains("smarttrackche.kwglobal.com")))
                    //    strPath = Strings.Replace(strPath, @"\" + strIntrnlJID + @"\", @"\" + strIntrnlJID + @"\" + strIntrnlID + @"\");




                    //if (SessionHandler.sOSName == "MAC" & blnReplaceMacAddress)
                    //{
                    //    strPath = strPath.Replace(@"\", "/");
                    //    strPath = "smb:" + strPath;
                    //}
                }
            }
            catch (Exception ex)
            {

            }

            return strPath;
        }


        private static bool Proc_Split_VolIss(string strIss, ref string strVolNo, ref string strIssNo)
        {
            try
            {
                if (strIss == "")
                {
                    strVolNo = "000";
                    strIssNo = "00";
                    return true;
                }

                strVolNo = Strings.Mid(strIss, 1, Strings.InStr(1, strIss, ":", Microsoft.VisualBasic.CompareMethod.Text) - 1);
                strIssNo = Strings.Mid(strIss, Strings.InStr(1, strIss, ":", Microsoft.VisualBasic.CompareMethod.Text) + 1);

                if (strVolNo.Length == 1)
                    strVolNo = "00" + strVolNo;
                if (strVolNo.Length == 2)
                    strVolNo = "0" + strVolNo;
                if (Strings.InStr(1, strIssNo, "-", Microsoft.VisualBasic.CompareMethod.Text) > 0)
                {
                    string strIssFirst = Strings.Mid(strIssNo, 1, Strings.InStr(1, strIssNo, "-", Microsoft.VisualBasic.CompareMethod.Text) - 1);
                    string strIssSec = Strings.Mid(strIssNo, Strings.InStr(1, strIssNo, "-", Microsoft.VisualBasic.CompareMethod.Text) + 1);
                    if (strIssFirst.Length == 1)
                        strIssFirst = "0" + strIssFirst;
                    if (strIssSec.Length == 1)
                        strIssSec = "0" + strIssSec;
                    strIssNo = strIssFirst + strIssSec;
                }
                else if (strIssNo.Length == 1)
                    strIssNo = "0" + strIssNo;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        
        public ActionResult Invoice()
        {
            ViewBag.getCustList = getCustList();
            ViewBag.getJList = getJournalList(convertToJSON: false);
            return View();
        }
        public dynamic getCustList()
        {
            try
            {
                string strQuery = "if exists (SELECT CM.CustSN FROM JBM_EmployeeConfig EC INNER JOIN JBM_CustomerMaster CM ON CM.CustID=EC.Description WHERE EC.ProcessID='CID0001' AND EC.EmpAutoID='" + Session["EmpAutoId"].ToString() + "' AND CM.CustType='" + Session["sCustAcc"].ToString() + "' AND CM.Cust_Disabled IS NULL) begin SELECT CM.CustSN,CM.CustID FROM JBM_EmployeeConfig EC INNER JOIN JBM_CustomerMaster CM ON CM.CustID=EC.Description WHERE EC.ProcessID='CID0001' AND EC.EmpAutoID='" + Session["EmpAutoId"].ToString() + "' AND CM.CustType='" + Session["sCustAcc"].ToString() + "' AND CM.Cust_Disabled IS NULL ORDER BY CM.CustSN DESC end else begin select c.CustSN,c.CustID from JBM_CustomerMaster c where c.Cust_Disabled is NULL and c.CustType='" + Session["sCustAcc"].ToString() + "' order by c.CustSn end";
                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                return from a in dt.AsEnumerable() select new SelectListItem { Text = a["CustSN"].ToString(), Value = a["CustID"].ToString() };                
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public ActionResult getInvoiceDetails(string jID, string process)
        {
            try
            {
                string ArticleInfo = "";
                if (Regex.IsMatch(Session["sCustAcc"].ToString().ToUpper(), "^(BK)$"))
                {
                    ArticleInfo = Session["sCustAcc"].ToString() + "_ChapterInfo";
                }
                else
                {
                    ArticleInfo = Session["sCustAcc"].ToString() + "_ArticleInfo";
                }

                string strQuery = "Select j.JBm_AutoID,J.JBm_Name,J.Jbm_Intrnl,j.JBM_ID,j.JBM_Intrnl,Isnull(j.JBM_PeName,'') as JBM_PeName,Isnull(j.JBM_Location,'') as JBM_Location,J.JBM_InvNo,J.JBM_InvDt,J.BM_DesiredPgCount as [TotalPgs],J.JBM_PeName,J.BM_Author,J.Title,M.Edition,M.HardBack_ISBN,M.PaperBack_ISBN,M.Ebook_ISBN,M.PONumber,J.DocketNo,M.BillingType ,j.JBM_Level,Case when M.JBM_PMOffshore='1' then (Case when J.JBM_Level='1' then 'ENPM11' when JBm_Level='2' then 'ENPM21'  when JBm_Level='3' then 'ENPM31' end) else (Case when J.JBM_Level='1' then 'ENPM12' when JBm_Level='2' then 'ENPM22'  when JBm_Level='3' then 'ENPM32' end) end as [PM_level_WBS], M.JBM_PMOffshore as  [PMReqOffshore] ,J.JBM_CeLevel , Case when J.JBM_CeRequiredOffshore='1' then (Case when J.JBM_CeLevel='0' then 'ENCE01' when J.JBM_CeLevel='1' then 'ENCE11' when J.JBM_CeLevel='2' then 'ENCE21'  when J.JBM_CeLevel='3' then 'ENCE31' end) else (Case when J.JBM_CeLevel='0' then 'ENCE02' when J.JBM_CeLevel='1' then 'ENCE12' when J.JBM_CeLevel='2' then 'ENCE22'  when J.JBM_CeLevel='3' then 'ENCE32' end) end as [CE_Level_WBS],J.JBM_CeRequiredOffshore ,J.JBM_ProofReadingOffshore,Case when J.JBM_ProofReadingOffshore='1' then 'ENPR01' else 'ENPR02' end as [ProofReading_WBS] ,Case when M.[IndexChrgTo]='Royalties' then IndexChrgTo when IndexChrgTo='Invoice' or IndexChrgTo='ISBN' then 'ISBN' else '' end as [IndexType],M.IndexChrgType as IndexSubType,case when J.JBM_IndexingOffshore='1' then 'ENSAI01' else 'ENSAI02' end as[Index_WBS],J.JBM_IndexingOffshore,M.[Index] ,M.LegalTabling as LegalTabling,'' As LegalTablingOffshore ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='simple') as Redraw_Simple ,case when (select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='simple')<>'0' then 'ENRS02'  else '' end as [Redraw_Simple_WBS] ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='Medium') as Redraw_Medium  ,case when (select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='Medium')<>'0' then 'ENRM02'  else '' end as [Redraw_Medium_WBS]  ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='Complex') as Redraw_Complex  ,case when (select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='yes' or F.FigReDraw='1') and ProcessComplexity='Complex')<>'0' then 'ENRC02'  else '' end as [Redraw_Complex_WBS]  ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Line Art') as [Line_Artwork] ,'ENLA02' as [Line_Artwork_WBS] ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='halfTones') as [Optimis_halftones] ,'ENOH02' as [Optimis_halftones_WBS] ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Relabel line art') as [Relabel_lineArt],'ENRA02' as [Relabel_lineArt_WBS] ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Grayscale') as [Grayscale],'ENGS02' as [Grayscale_WBS] ,(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Color') as [4Color],'EN4C02' as [4Color_WBS],J.ArtOffshore as [ArtEstimateOffshore] ,M.CoverPpcNumber as [Template_COV_Create],CoverPageType as [Template_CoverpPage_type],'ENTC02' as [Template_COVCreate_WBS],M.Typesetting_level, Case when M.typesetting_level='1' then 'ENTL12' when M.typesetting_level='2' then 'ENTL22'  when M.typesetting_level='3' then 'ENTL32' end as [TSP_Level],M.IsNewTemplate,M.JBM_TemplateOffshore,Case When M.IsNewTemplate='1'and M.JBM_TemplateOffshore='1' then 'ENCT01' when M.IsNewTemplate='1' and  (M.JBM_TemplateOffshore is null or M.JBM_TemplateOffshore='0') then 'ENCT02' end as [CreateTemplateFromScratch_WBS],Isnull(J.eBooks,0) as eBooks, Case When (J.eBooks='1') and (M.[XML]='1' ) then 'ENEXML02' end as [ebooks_XML_WBS] , case when (J.eBooks='1') and (M.Print_PDF='1' ) then 'ENEPDF02'end as [ebooks_PDF_WBS] ,M.Rekeytxt,Case When M.Rekeytxt is not null then 'ENRT02' else '' end [Rekeytxt_WBS] ,M.RekeyEquation, Case When M.RekeyEquation is not null then 'ENRT02' else '' end [RekeyEquation_WBS],(select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Alt Text')  as AltText,Case when AltTextOffshore ='1' and (select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Alt Text')<>0 then 'ENAW01' when (select Count(*) from " + Session["sCustAcc"].ToString() + "_FigDetails F Inner join " + ArticleInfo + " A On A.AutoartID=F.AutoArtID and A.Wip=1 and A.JBM_AutoID=M.JBM_AutoID and (F.FigReDraw='no' or F.FigReDraw='0' or F.FigReDraw is null) and FigureMode='Alt Text')<>0 then 'ENAW02' else '' end [AltText_WBS],J.AltTextOffshore,J.BM_NofFigs,M.WordCount,M.WIPclosed_Date,M.ALtTextPgs,Isnull(M.PM_PageCount,J.BM_DesiredPgCount) as PM_PageCount,Isnull(M.TypeSetting_PageCount,J.BM_DesiredPgCount) as TypeSetting_PageCount,Isnull(M.eBookCreation_PageCount,J.BM_DesiredPgCount) as eBookCreation_PageCount,Isnull(M.CE_WordCount,M.WordCount) as CE_WordCount ,Isnull(M.PR_WordCount,M.WordCount) as PR_WordCount,Isnull(M.Index_WordCount,M.WordCount) as Index_WordCount ,M.Proofread from JBm_info J Inner Join " + Session["sCustAcc"].ToString() + "_ProjectManagement M On J.JBM_AutoID=M.JBM_AutoID  where M.JBm_AutoID in ('" + jID + "') and   M.WIPclosed_Date is Null; Select PM.ProcessID,PM.Process,P.Description,P.Rate,P.Stage,P.ShortDescript,P.SubProcess,(select Process from JBM_ProcessMaster k where K.ProcessID=P.SubProcess) as [Subprocess Desc] from " + Session["sCustAcc"].ToString() + "_ProcessInfo P , JBM_ProcessMaster PM where P.ProcessID=PM.ProcessID  and Description like 'Core%' and (select Process from JBM_ProcessMaster k where K.ProcessID=P.SubProcess) like '" + process + "%' Order by PM.ProcessID,P.SubProcess ";

                DataSet ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                var jsonResult = Json(new { result = "Success", data = (ds.Tables[0].Rows.Count > 0) ? JsonConvert.SerializeObject(ds.Tables[0]) : "", processlist = (ds.Tables[1].Rows.Count > 0) ? JsonConvert.SerializeObject(ds.Tables[1]) : "" }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;
                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    result = "Error",
                    data = ex.Message
                }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult saveInvoiceDetails(string Author, string Title, string ISBN1, string ISBN2, string ISBN3, string PONumber, string DocketNo, string isFinal, string PMLevel, string PMReqOffshore, string CELevel, string CEReqOffshore, string LTReqOffshore, string PRReqOffshore, string IndexReqOffshore, string Template_ReqOffshore, string TSlevel, string AltText_ReqOffshore, string JBM_AutoID, string InvoiceNo, string InvoiceDate, string Index, string CoverType, string Rekeytxt, string RekeyEquation, string BillingType, string eBooks, string LegalTablingOffshore, string ALtTextPgs, string TotalPages, string WordCount, string htmlContent, string xmlContent, string pdfFileName, string NofFigs, string LegalTabling, string PMPageCount, string CEWordCount, string PRWordCount, string IndexWordCount, string TSPageCount, string EBookPageCount)
        {
            try
            {
                string strQuery = "select 1 as RN,JBM_AutoID,BM_Author,Title, DocketNo,JBM_Level, JBM_CeLevel, JBM_CeRequiredOffshore, JBM_ProofReadingOffshore, JBM_IndexingOffshore, AltTextOffshore, JBM_InvNo, JBM_InvDt, eBooks, BM_DesiredPgCount, BM_NofFigs into #temp1 from JBM_Info where JBM_AutoID = '" + JBM_AutoID + "';";

                strQuery = strQuery + "Update JBM_Info set BM_Author ='" + Author + "', Title='" + Title + "', DocketNo='" + DocketNo + "',JBM_Level=Nullif('" + PMLevel + "',''),JBM_CeLevel=Nullif('" + CELevel + "',''),JBM_CeRequiredOffshore=Nullif('" + CEReqOffshore + "','2'),JBM_ProofReadingOffshore=Nullif('" + PRReqOffshore + "','2'),JBM_IndexingOffshore=Nullif('" + IndexReqOffshore + "','2'),AltTextOffshore=Nullif('" + AltText_ReqOffshore + "','2'),JBM_InvNo='" + InvoiceNo + "',JBM_InvDt=Nullif('" + InvoiceDate + "',''),eBooks=Nullif('" + eBooks + "',''),BM_DesiredPgCount='" + TotalPages + "',BM_NofFigs='" + NofFigs + "' where JBM_AutoID='" + JBM_AutoID + "'; ";

                strQuery = strQuery + "if @@ROWCOUNT >=1 begin insert into #temp1 select 2 as RN,JBM_AutoID,BM_Author,Title, DocketNo,JBM_Level, JBM_CeLevel, JBM_CeRequiredOffshore, JBM_ProofReadingOffshore, JBM_IndexingOffshore, AltTextOffshore, JBM_InvNo, JBM_InvDt, eBooks, BM_DesiredPgCount, BM_NofFigs from JBM_Info where JBM_AutoID = '" + JBM_AutoID + "' end;";

                strQuery = strQuery + "select 1 as RN, HardBack_ISBN, PaperBack_ISBN, Ebook_ISBN, PONumber, JBM_PMOffshore, JBM_TemplateOffshore, Typesetting_level, IndexChrgTo, CoverPageType, Rekeytxt, RekeyEquation, BillingType, ALtTextPgs, WordCount,LegalTabling,PM_PageCount,CE_WordCount,PR_WordCount,Index_WordCount,TypeSetting_PageCount,eBookCreation_PageCount into #temp2 from  " + Session["sCustAcc"].ToString() + "_ProjectManagement where JBM_AutoID = '" + JBM_AutoID + "';";

                strQuery = strQuery + "Update " + Session["sCustAcc"].ToString() + "_ProjectManagement set HardBack_ISBN='" + ISBN1 + "',PaperBack_ISBN='" + ISBN2 + "',Ebook_ISBN='" + ISBN3 + "',PONumber='" + PONumber + "',JBM_PMOffshore=Nullif('" + PMReqOffshore + "','2'),JBM_TemplateOffshore=Nullif('" + Template_ReqOffshore + "','2'),Typesetting_level=Nullif('" + TSlevel + "',''),IndexChrgTo=Nullif('" + Index + "',''),CoverPageType=Nullif('" + CoverType + "',''),Rekeytxt='" + Rekeytxt + "',RekeyEquation='" + RekeyEquation + "', BillingType='" + BillingType + "',ALtTextPgs='" + ALtTextPgs + "',WordCount='" + WordCount + "',LegalTabling='" + LegalTabling + "',PM_PageCount=Nullif('" + PMPageCount + "',''),CE_WordCount=Nullif('" + CEWordCount + "',''),PR_WordCount=Nullif('" + PRWordCount + "',''),Index_WordCount=Nullif('" + IndexWordCount + "',''),TypeSetting_PageCount=Nullif('" + TSPageCount + "',''),eBookCreation_PageCount=Nullif('" + EBookPageCount + "','') where JBM_AutoID='" + JBM_AutoID + "';";

                strQuery = strQuery + "if @@ROWCOUNT >=1 begin insert into #temp2 select 2 as RN, HardBack_ISBN, PaperBack_ISBN, Ebook_ISBN, PONumber, JBM_PMOffshore, JBM_TemplateOffshore, Typesetting_level, IndexChrgTo, CoverPageType, Rekeytxt, RekeyEquation, BillingType, ALtTextPgs, WordCount,LegalTabling,PM_PageCount,CE_WordCount,PR_WordCount,Index_WordCount,TypeSetting_PageCount,eBookCreation_PageCount from  " + Session["sCustAcc"].ToString() + "_ProjectManagement where JBM_AutoID = '" + JBM_AutoID + "' end;";

                strQuery = strQuery + "select JBM_AutoID,BM_Author,Title, DocketNo,JBM_Level, JBM_CeLevel, JBM_CeRequiredOffshore, JBM_ProofReadingOffshore, JBM_IndexingOffshore, AltTextOffshore, JBM_InvNo, JBM_InvDt, eBooks, BM_DesiredPgCount, BM_NofFigs from #temp1 order by RN;select HardBack_ISBN, PaperBack_ISBN, Ebook_ISBN, PONumber, JBM_PMOffshore, JBM_TemplateOffshore, Typesetting_level, IndexChrgTo, CoverPageType, Rekeytxt, RekeyEquation, BillingType, ALtTextPgs, WordCount,LegalTabling,PM_PageCount,CE_WordCount,PR_WordCount,Index_WordCount,TypeSetting_PageCount,eBookCreation_PageCount from #temp2 order by RN; drop table #temp1;drop table #temp2";

                //bool result = DBProc.UpdateRecord(strQuery, Session["sConnSiteDB"].ToString());
                DataSet ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                if (ds.Tables.Count == 2)
                {
                    string strUpdatedColumns = "";

                    foreach (DataTable dt in ds.Tables)
                    {
                        if (dt.Rows.Count == 2)
                        {
                            foreach (DataColumn dc in dt.Columns)
                            {
                                if (dt.Rows[0][dc.ColumnName].ToString() != dt.Rows[1][dc.ColumnName].ToString())
                                {
                                    strUpdatedColumns = strUpdatedColumns + "<tr><td>" + dc.ColumnName + "</td>" + "<td>" + dt.Rows[0][dc.ColumnName] + "</td>" + "<td>" + dt.Rows[1][dc.ColumnName] + "</td></tr>";
                                }
                            }
                        }
                    }


                    RL.ArtDet objArtDet = new RL.ArtDet();
                    objArtDet.AD.CustAccess = Session["sCustAcc"].ToString();
                    objArtDet.AD.ConnString = DBProc.getConnection(Session["sConnSiteDB"].ToString()).ConnectionString;

                    string InvoicePath = RL.clsFileIO.Proc_Get_Directory_Path(ref objArtDet, "F5000-Bill", false);

                    if (InvoicePath != "")
                    {
                        InvoicePath = Path.Combine(InvoicePath, "GenerateInvoice");
                    }
                    else
                    {
                        return Json("False");
                    }

                    if (strUpdatedColumns != "")
                    {
                        string trackingPath = Path.Combine(InvoicePath, "Tracking", JBM_AutoID);

                        if (!Directory.Exists(trackingPath))
                        {
                            Directory.CreateDirectory(trackingPath);
                        }

                        System.IO.File.AppendAllText(Path.Combine(trackingPath, pdfFileName + ".html"), strUpdatedColumns);
                    }

                    if (htmlContent != "")
                    {
                        xmlContent = xmlContent.Replace("&nbsp;", " ");
                        System.Xml.Linq.XElement xEle = System.Xml.Linq.XElement.Parse(xmlContent);
                        string htmlPath = Path.Combine(InvoicePath, "HTML");
                        string pdfPath = Path.Combine(InvoicePath, "PDF");
                        string xmlPath = Path.Combine(InvoicePath, "XML");
                        string zipPath = Path.Combine(InvoicePath, "Zip");
                        string htmlName = pdfFileName;

                        if (!Directory.Exists(htmlPath))
                        {
                            Directory.CreateDirectory(htmlPath);
                        }

                        htmlPath = Path.Combine(htmlPath, pdfFileName + ".html");

                        if (!Directory.Exists(pdfPath))
                        {
                            Directory.CreateDirectory(pdfPath);
                        }

                        pdfPath = Path.Combine(pdfPath, pdfFileName + ".pdf");

                        if (!Directory.Exists(xmlPath))
                        {
                            Directory.CreateDirectory(xmlPath);
                        }

                        xmlPath = Path.Combine(xmlPath, pdfFileName + ".xml");

                        if (!Directory.Exists(zipPath))
                        {
                            Directory.CreateDirectory(zipPath);
                        }

                        zipPath = Path.Combine(zipPath, pdfFileName + ".zip");

                        System.IO.File.AppendAllText(htmlPath, htmlContent);
                        System.IO.File.AppendAllText(xmlPath, xEle.ToString());

                        ProcessStartInfo startInfo = new ProcessStartInfo();

                        startInfo.CreateNoWindow = false;
                        startInfo.UseShellExecute = false;
                        startInfo.FileName = @"C:\Applications\ChromeHtmltoPdf_265\ChromeHtmlToPdfConsole.exe";// @"C:\Applications\GenerateInvoice\exe\wkhtmltopdf.exe";
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        //startInfo.Arguments = Path.Combine(htmlPath, pdfFileName + ".html") + " " + Path.Combine(pdfPath, pdfFileName + ".pdf");
                        startInfo.Arguments = "--input " + htmlPath + " --output " + pdfPath;

                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.WaitForExit();
                        }

                        Ionic.Zip.ZipEntry zEntry = new Ionic.Zip.ZipEntry();

                        using (Ionic.Zip.ZipFile zFile = new Ionic.Zip.ZipFile())
                        {
                            if (System.IO.File.Exists(pdfPath))
                            {
                                //zFile.AddEntry(pdfPath, fileName);
                                zFile.AddFile(pdfPath, pdfFileName + ".pdf");
                            }

                            if (System.IO.File.Exists(xmlPath))
                            {
                                //zFile.AddEntry(xmlPath, fileName);
                                zFile.AddFile(xmlPath, pdfFileName + ".xml");
                            }
                            zFile.Save(zipPath);
                            TempData["zipPath"] = zipPath;
                        }

                        return Json("True");
                    }
                    else
                    {
                        return Json("True");
                    }
                }
                else
                {
                    return Json("False");
                }

            }
            catch (Exception ex)
            {
                return Json("False");
            }
        }

        public void downloadInvoice(string fileName)
        {
            try
            {           
                string zipPath = TempData["zipPath"].ToString();
                Response.Clear();
                Response.AddHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(zipPath));
                Response.ContentType = "application/zip";
                Response.WriteFile(zipPath);
                Response.End();

            }
            catch (Exception ex)
            {

            }
        }

        public string invoiceFTPUpload(string JBM_AutoID,string fileName,string InvoiceNo)
        {
            string result = "";
            try
            {
                RL.ArtDet objArtDet = new RL.ArtDet();
                objArtDet.AD.CustAccess = Session["sCustAcc"].ToString();
                objArtDet.AD.ConnString = DBProc.getConnection(Session["sConnSiteDB"].ToString()).ConnectionString;

                string InvoicePath = RL.clsFileIO.Proc_Get_Directory_Path(ref objArtDet, "F5000-Bill", false);
                string zipPath = Path.Combine(InvoicePath, "GenerateInvoice", "Zip", fileName + ".zip");
                string hostName = "";
                string username = "";
                string pwd = "";
                string wrkDirectory = "";
                int port = 2222;
                string strMailFrom = "";
                string strMailTo = "";
                string strMailCC = "";
                string strMailBCC = "";
                string strMailSubject = "";
                string strMailBody = "";

                string strQry = "select top 1 ftpHost,ftpUID,ftpPWD,ftpPath from JBM_Ftp_Details where Stage='Billing' and CustAccess='" + Session["sCustAcc"].ToString() + "';";
                strQry += "select top 1 MailFrom,MailTo,MailCc,MailBcc,MsgBody,MsgSubject from JBM_FTP_Stage_Details a inner join JBM_messageInfo b on b.MsgID = a.MailSubject where Stage = 'Billing' and CustType='" + Session["sCustAcc"].ToString() + "';";
                strQry += "select JBM_ID,DateTimeStamp as [is_exists] from JBM_Info i left join JBM_FtpUp f on f.JID=i.JBM_AutoID and f.Stage='Billing' where JBM_AutoID='" + JBM_AutoID + "'";
                
                DataSet ds = DBProc.GetResultasDataSet(strQry, Session["sConnSiteDB"].ToString());

                if (ds.Tables[0].Rows.Count > 0)
                {
                    hostName = ds.Tables[0].Rows[0]["ftpHost"].ToString();
                    username = ds.Tables[0].Rows[0]["ftpUID"].ToString();
                    pwd = ds.Tables[0].Rows[0]["ftpPWD"].ToString();
                    wrkDirectory = ds.Tables[0].Rows[0]["ftpPath"].ToString();
                    port = 2222;
                }
                else
                {
                    return "FTPDetailsMissing";
                }

                // Commented by Sorimuthu for test
                //if (ds.Tables[2].Rows[0]["is_exists"].ToString() != "")
                //{
                //    return "AlreadyUploaded";
                //}

                if (ds.Tables[1].Rows.Count > 0)
                {
                    strMailFrom = ds.Tables[1].Rows[0]["MailFrom"].ToString().Replace("'","''");
                    strMailTo = ds.Tables[1].Rows[0]["MailTo"].ToString().Replace("'", "''");
                    strMailCC = ds.Tables[1].Rows[0]["MailCc"].ToString().Replace("'", "''");
                    strMailBCC = ds.Tables[1].Rows[0]["MailBcc"].ToString().Replace("'", "''");
                    strMailSubject = ds.Tables[1].Rows[0]["MsgSubject"].ToString().Replace("###Project###", ds.Tables[2].Rows[0]["JBM_ID"].ToString()).Replace("###InvoiceNo###", InvoiceNo);
                    strMailBody = ds.Tables[1].Rows[0]["MsgBody"].ToString().Replace("###Project###", ds.Tables[2].Rows[0]["JBM_ID"].ToString()).Replace("###InvoiceNo###", InvoiceNo);
                }
                else
                {
                    return "MailDetailsMissing";
                }

                // Commented by Sorimuthu for test
                //using (SftpClient sftpClient = new SftpClient(hostName, port, username, pwd))
                //{
                //    Console.WriteLine("Connect to the SFTP server");
                //    sftpClient.Connect();
                //    sftpClient.ChangeDirectory(wrkDirectory);
                //    Console.WriteLine("Creating FileStream object to stream a file");
                //    //Upload file
                //    using (FileStream fs = new FileStream(zipPath, FileMode.Open))
                //    {
                //        sftpClient.BufferSize = 4 * 1024;
                //        sftpClient.UploadFile(fs, Path.GetFileName(zipPath));
                //    }                   
                //    sftpClient.Dispose();
                //}

                string strQuery = "insert into JBM_MailInfo (EmpAutoID,JBM_AutoID,MailEventID,MailStatus,MailInitDate,MailFrom,MailTo,MailCc,MailBCc,MailSub,MailBody,IsBodyHtml,MailPriority,MailType,CustType) values('" + Session["EmpAutoId"].ToString() + "','" + JBM_AutoID + "','INV001',0,GETDATE(),'" + strMailFrom + "','" + strMailTo + "','" + strMailCC + "','" + strMailBCC + "','" + strMailSubject + "','" + strMailBody + "',1,2,'Direct','" + Session["sCustAcc"].ToString() + "');";

                strQuery += "insert into JBM_FtpUp (EmpAutoId,AutoArtId,IntrnlID,JID,FTPSite,FileName,DtStamp,Stage,CustType,DateTimeStamp,Wf_NextStage,InProcStatus) values('" + Session["EmpAutoId"].ToString() + "','','','" + JBM_AutoID + "','" + hostName + wrkDirectory + "','',Getdate(),'Billing','" + Session["sCustAcc"].ToString() + "',Getdate(),'',1);";
                result = DBProc.InsertRecord(strQuery, Session["sConnSiteDB"].ToString());
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// Note always keep in end of the controller
        /// </summary>
        /// <param name="filterContext"></param>
        protected override void OnException(ExceptionContext filterContext)
        {
            filterContext.ExceptionHandled = true;

            //Log the error!!

            //Redirect to action
            filterContext.Result = RedirectToAction("ErrorHandler", "Index");
            Session["sErrorMessage"] = filterContext.Exception.Message;
            // OR return specific view
            filterContext.Result = new ViewResult
            {
                ViewName = "~/Views/ErrorHandler/Index.cshtml"
            };
        }

        public ActionResult Dashboard()
        {
            return View();
        }
        public ActionResult getTATEventData(string DateRangeFrom, string DateRangeTo, string TeamID, string SubTeamID, string[] JournalList)
        {
            try
            {
                string strHolidaysList = string.Empty;
                DataSet dsHoliday = new DataSet();
                dsHoliday = DBProc.GetResultasDataSet("SELECT (STUFF((SELECT '{' + CONVERT(varchar,HolidayDate,110) + '}' as 'DATE' FROM tbl_HolidayList FOR XML PATH('')), 1, 0, '')) AS [HoliDays]", Session["sConnSiteDB"].ToString());
                if (dsHoliday.Tables[0].Rows.Count > 0)
                {
                    strHolidaysList = dsHoliday.Tables[0].Rows[0][0].ToString();
                    strHolidaysList = strHolidaysList.Replace("<DATE>", "");
                    strHolidaysList = strHolidaysList.Replace("</DATE>", "");
                }

                string strArtStageQuery = string.Empty;
                List<string> artStageTypeList = clsCollec.lstFPArtStage();

                for (int i = 0; i < artStageTypeList.Count; i++)
                {
                    if (i == artStageTypeList.Count - 1)
                    {
                        strArtStageQuery += "s.artstagetypeid like '%" + artStageTypeList[i].Trim() + "%' ";
                    }
                    else
                    {
                        strArtStageQuery += "s.artstagetypeid like '%" + artStageTypeList[i].Trim() + "%' or ";
                    }
                }

                clsINI obj = new clsINI();
                string strJBMIDs = string.Empty;
                string tempJID = string.Empty;
                if (JournalList != null)
                {
                    foreach (string strJID in JournalList)
                    {
                        if (tempJID != "")
                        {
                            tempJID = tempJID + ",";
                        }
                        tempJID = tempJID + obj.getJournlIDByTeam("TF", strJID, Session["sConnSiteDB"].ToString());
                    }
                    strJBMIDs = " a.JBM_AutoID in (" + tempJID + ") and ";
                }

                // For pilot journals
                string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                XmlNodeList objNodelist;
                string strConnValue = string.Empty;
                XmlDocument objxml = new XmlDocument();
                objxml.Load(strPath);
                if (objxml.InnerXml != "Nothing")
                {
                    objNodelist = objxml.SelectNodes("//config/WMSPilotJournals");
                    if (objNodelist.Count > 0)
                    {
                        strJBMIDs = " a.JBM_AutoID in ('" + objNodelist.Item(0).InnerText.ToString().Replace("|", "','") + "') and ";
                    }
                }


                try
                {
                    DataTable dtProArtStaDes = DBProc.GetResultasDataTbl("Select DeptActivity,StageDesc,Stage from " + Init_Tables.gTblProdArtStatusDesDept + "", Session["sConnSiteDB"].ToString());
                    DataTable dtEmpDet = DBProc.GetResultasDataTbl("Select EmpLogin,EmpAutoID,EmpName,EmpSurname from " + Init_Tables.gTblEmployee + " ", Session["sConnSiteDB"].ToString());
                    //(CASE WHEN p.Process = 'Waiting for ZIP File Movement' THEN 'ZIP' WHEN p.Process = 'Waiting for Smart Track Entry' THEN 'Wait' WHEN  p.Process ='Database update Success' THEN 'Database' ELSE p.Process END  )
                    string strQueryFinal = "DECLARE @TotMinss INT=0; Select [JournalID],[SmartTrack_ID],ArticleID,[Received_Date],[DueDate],p.Process as [Current_Activity],(convert(varchar, datediff (s, [DueDate], getdate()) / (60 * 60 * 24)) + ' day(s) ' + convert(varchar, dateadd(s, datediff (s, [DueDate], getdate()), convert(datetime2, '0001-01-01')), 108)) as [Delayed_Hrs],'' as [Revised_Hrs], '' as [Action],JBM_AutoID,[ESA_Status],[TotalMins],p.Process as [ActvityName],[CEOnShore] from (Select ROW_NUMBER() OVER (PARTITION BY eea.JBM_AutoID, eea.ArticleID ORDER BY  eea.ExternalEventActDate DESC) AS [RN], eea.JBM_AutoID as [JournalID], eea.AutoArtID as [SmartTrack_ID],eea.ArticleID, eea.ExternalEventActDate  as [Received_Date], isnull(eea.ExternalEventTriggerd,'') as [DueDate],format(isnull(CONVERT(INT,eea.Status),0),'EVT000') as [Status],'' as [Delayed_Hrs],'' as [Revised_Hrs], '' as [Action],j.JBM_AutoID, '' as [ESA_Status], DATEDIFF(MINUTE, isnull(eea.ExternalEventTriggerd,GETDATE()), GETDATE()) as [TotalMins],'' as [CEOnShore],'' as [ActvityName]  from tf_externaleventaccess eea join jbm_info j on j.JBM_Intrnl=eea.JBM_AutoID where eea.ExternalEventID='EV003' and eea.Status in (400,4,5,7,8,6,11) and eea.Stage in ('Typesetting','Tagging')) as e left join JBM_ProcessMaster p on e.Status=p.ProcessID Where RN=1 and [TotalMins] > 120 and Received_Date between cast(convert(varchar(50), '" + DateRangeFrom + "',103) as datetime) and cast(convert(varchar(50), '" + DateRangeTo + "',103) as datetime) order by [TotalMins] desc";

                    DataTable dtWip = new DataTable();
                    dtWip = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());

                    //string strStageInfoQuery = "Select [JournalID],[SmartTrack_ID],ArticleID,[Received_Date],[DueDate],[FilterColumn] as [Current_Activity],[Delayed_Hrs],[Revised_Hrs],[Action],JBM_AutoID,[ESA_Status],[TotalMins],[ActvityName],[CEOnShore],ExternalEventActDate,ExternalEventReplay,ArtStageTypeID,CEDispDate from (Select *, CASE WHEN ([ActvityName] = 'Login' OR [ActvityName] = 'Project Management')  and TotalMins> 60 THEN 'Login' WHEN ([ActvityName] = 'CleanUP' OR [ActvityName] = 'Auto-CI')  and TotalMins> 120 THEN 'CleanUp' WHEN [ActvityName] = 'Pre-Edit'   and TotalMins> 420 THEN [ActvityName] WHEN ([ActvityName] = 'Normalization' OR [ActvityName] = 'Auto-ESA' OR [ActvityName] = 'Mechanical Editing')  and TotalMins> 480 THEN 'Normalization' WHEN [ActvityName] = 'CE' AND ESA_Status ='B' and TotalMins> 1920 THEN [ActvityName] WHEN [ActvityName] = 'CE' AND ESA_Status ='C' and  CEOnShore is null and TotalMins> 5760 THEN [ActvityName] WHEN [ActvityName] = 'CE' AND ESA_Status ='C' and  CEOnShore=1 and TotalMins> 8640 THEN [ActvityName] WHEN [ActvityName] = 'ML' AND ESA_Status ='A' and TotalMins> 840 THEN [ActvityName] WHEN [ActvityName] = 'ML' AND ESA_Status ='B' and TotalMins> 2280 THEN [ActvityName] WHEN [ActvityName] = 'ML' AND ESA_Status ='C' and  CEOnShore is null and TotalMins> 6120 THEN [ActvityName] WHEN [ActvityName] = 'ML' AND ESA_Status ='C' and  CEOnShore=1 and TotalMins> 9000 THEN [ActvityName] WHEN [ActvityName] = 'Pag' AND ESA_Status ='A' and TotalMins> 1200 THEN [ActvityName] WHEN [ActvityName] = 'Pag' AND ESA_Status ='B' and TotalMins> 2640 THEN [ActvityName] WHEN [ActvityName] = 'Pag' AND ESA_Status ='C' and  CEOnShore is null and TotalMins> 6480 THEN [ActvityName] WHEN [ActvityName] = 'Pag' AND ESA_Status ='C' and  CEOnShore=1 and TotalMins> 9360 THEN [ActvityName] WHEN [ActvityName] = 'TE' AND ESA_Status ='A' and TotalMins> 1440 THEN [ActvityName] WHEN [ActvityName] = 'TE' AND ESA_Status ='B' and TotalMins> 2880 THEN [ActvityName] WHEN [ActvityName] = 'TE' AND ESA_Status ='C' and  CEOnShore is null and TotalMins> 6720 THEN [ActvityName] WHEN [ActvityName] = 'TE' AND ESA_Status ='C' and  CEOnShore=1 and TotalMins> 9600 THEN [ActvityName] END as [FilterColumn] from (Select ROW_NUMBER() OVER (PARTITION BY JournalID,ArticleID ORDER BY  Revised_Hrs DESC) AS [RNo],* from (Select * from (Select j.JBM_Intrnl as [JournalID],s.AutoArtID as [SmartTrack_ID], a.ChapterID as[ArticleID],s.ReceivedDate as [Received_Date],isnull(s.DueDate,'') as[DueDate],s.ArtStageTypeID as [Current_Activity],dbo.fn_CurrentStage(s.ArtStageTypeID) as [ActvityName],(convert(varchar, datediff (s, s.ReceivedDate, getdate()) / (60 * 60 * 24)) + ' day(s) ' + convert(varchar, dateadd(s, datediff (s, s.ReceivedDate, getdate()), convert(datetime2, '0001-01-01')), 108)) as [Delayed_Hrs],'' as [Revised_Hrs], '' as [Action], a.Jbm_AutoID,substring(isnull(s.ESA_Status,'0'),1,1) as [ESA_Status],DATEDIFF(MINUTE, isnull(s.ReceivedDate,GETDATE()), GETDATE()) as [TotalMins],j.JBM_CeRequiredOffShore as [CEOnShore],a.ExternalEventActDate,a.ExternalEventReplay,s.ArtStageTypeID as [ArtStageTypeID],s.CEDispDate as [CEDispDate] from tf_stageinfo s left join tf_articleinfo a on s.autoartid=a.autoartid join jbm_info j on j.jbm_autoid=a.jbm_autoid where " + strJBMIDs + " (" + strArtStageQuery + ") and s.RevFinStage='FP' and s.DispatchDate is null and s.ReceivedDate between cast(convert(varchar(50), '" + DateRangeFrom + "',103) as datetime) and cast(convert(varchar(50), '" + DateRangeTo + "',103) as datetime)) as t WHERE [TotalMins] > 60 ";

                    string strStageInfoQuery = "Select [JournalID],[SmartTrack_ID],ArticleID,[Received_Date],[DueDate],[FilterColumn] as [Current_Activity],[Delayed_Hrs],[Revised_Hrs],[Action],JBM_AutoID,[ESA_Status],[TotalMins],[ActvityName],[CEOnShore],ExternalEventActDate,ExternalEventReplay,ArtStageTypeID,CEDispDate from (Select *, CASE WHEN ([ActvityName] = 'Login' OR [ActvityName] = 'Project Management')  THEN 'Login' WHEN ([ActvityName] = 'CleanUP' OR [ActvityName] = 'Auto-CI') THEN 'CleanUp' WHEN ([ActvityName] = 'Normalization' OR [ActvityName] = 'Auto-ESA' OR [ActvityName] = 'Mechanical Editing') THEN 'Normalization' WHEN [ActvityName]='Auto-Pagination' THEN 'Pag' ELSE [ActvityName] END as [FilterColumn] from (Select ROW_NUMBER() OVER (PARTITION BY JournalID,ArticleID ORDER BY  Revised_Hrs DESC) AS [RNo],* from (Select * from (Select j.JBM_Intrnl as [JournalID],s.AutoArtID as [SmartTrack_ID], a.ChapterID as[ArticleID],s.ReceivedDate as [Received_Date],isnull(s.DueDate,'') as[DueDate],s.ArtStageTypeID as [Current_Activity],dbo.fn_CurrentStage(s.ArtStageTypeID) as [ActvityName],(convert(varchar, datediff (s, s.ReceivedDate, getdate()) / (60 * 60 * 24)) + ' day(s) ' + convert(varchar, dateadd(s, datediff (s, s.ReceivedDate, getdate()), convert(datetime2, '0001-01-01')), 108)) as [Delayed_Hrs],'' as [Revised_Hrs], '' as [Action], a.Jbm_AutoID,substring(isnull(s.ESA_Status,'0'),1,1) as [ESA_Status],DATEDIFF(MINUTE, isnull(s.ReceivedDate,GETDATE()), GETDATE()) as [TotalMins],j.JBM_CeRequiredOffShore as [CEOnShore],a.ExternalEventActDate,a.ExternalEventReplay,s.ArtStageTypeID as [ArtStageTypeID],s.CEDispDate as [CEDispDate] from tf_stageinfo s left join tf_articleinfo a on s.autoartid=a.autoartid join jbm_info j on j.jbm_autoid=a.jbm_autoid where " + strJBMIDs + " (" + strArtStageQuery + ") and s.RevFinStage='FP' and s.DispatchDate is null and s.ReceivedDate between cast(convert(varchar(50), '" + DateRangeFrom + "',103) as datetime) and cast(convert(varchar(50), '" + DateRangeTo + "',103) as datetime)) as t WHERE [TotalMins] > 60 ";

                    strStageInfoQuery += " UNION Select j.JBM_Intrnl as [JournalID],s.AutoArtID as [SmartTrack_ID], a.ChapterID as[ArticleID],s.ReceivedDate as [Received_Date],isnull(s.DueDate,'') as[DueDate],s.ArtStageTypeID as [Current_Activity],CASE WHEN dbo.fn_CurrentStage(s.ArtStageTypeID) ='Pass2FreeLancer' THEN 'CE' ELSE dbo.fn_CurrentStage(s.ArtStageTypeID)  END  as [ActvityName],(convert(varchar, datediff (s, s.ReceivedDate, getdate()) / (60 * 60 * 24)) + ' day(s) ' + convert(varchar, dateadd(s, datediff (s, s.ReceivedDate, getdate()), convert(datetime2, '0001-01-01')), 108)) as [Delayed_Hrs],p.ShortDescript as [Revised_Hrs], '' as [Action], a.Jbm_AutoID,Substring(isnull(s.ESA_Status,'0'),1,1) as [ESA_Status],DATEDIFF(MINUTE, isnull(s.ReceivedDate,GETDATE()), GETDATE()) as [TotalMins],j.JBM_CeRequiredOffShore as [CEOnShore],a.ExternalEventActDate as ExternalEventActDate, a.ExternalEventReplay as ExternalEventReplay,s.ArtStageTypeID as [ArtStageTypeID],s.CEDispDate as [CEDispDate] from TF_Processinfo p left join tf_stageinfo s on s.AutoartID=p.AutoArtID join tf_articleinfo a on p.autoartid=a.autoartid join jbm_info j on j.jbm_autoid=a.jbm_autoid  where p.AutoartID=s.AutoArtID and s.DispatchDate is null and s.RevFinStage='FP' and p.Stage='FP' and p.ProcessID='EVT014' and p.Status=0 and s.ReceivedDate between cast(convert(varchar(50), '" + DateRangeFrom + "',103) as datetime) and cast(convert(varchar(50), '" + DateRangeTo + "',103) as datetime)) as k) as t WHERE RNo=1) as m WHERE FilterColumn is not null order by [Delayed_Hrs] desc";

                    DataTable dtStageWip = new DataTable();
                    dtStageWip = DBProc.GetResultasDataTbl(strStageInfoQuery, Session["sConnSiteDB"].ToString());

                    // Join the two datatables
                    dtWip.Merge(dtStageWip);

                    DataRow dr = dtWip.AsEnumerable().First();  
                    foreach (DataRow row in dtWip.Rows)
                    {
                        try
                        {
                            // Round Off Hours
                            //DateTime dateTime = new DateTime(2023, 05, 16, 8, 31, 12);
                            //var updated = dateTime.AddMinutes(30);
                            DateTime start = DateTime.Parse(row["Received_Date"].ToString());
                            DateTime dueDate = DateTime.Parse(row["DueDate"].ToString());
      
                            DateTime end = DateTime.Now;

                            if (row["SmartTrack_ID"].ToString() == "T1014520")
                            {
                                Console.WriteLine("T1012687");
                            }

                            string strCEDispDate = row["CEDispDate"].ToString();
                            string strArtStageType = row["ArtStageTypeID"].ToString();
                            // collecting Saturday, sunday and holidays hrs  // Saturday also excluded Devendra discussed with Mallik and confirmed.
                            int intHolidayHrs = 0;

                            for (var i = start; i < end; i = i.AddHours(1))
                            {
                                if (i.DayOfWeek == DayOfWeek.Saturday | i.DayOfWeek == DayOfWeek.Sunday | (strHolidaysList.Contains("{" + i.Month.ToString("D2") + "-" + i.Day.ToString("D2") + "-" + i.Year + "}")))
                                {
                                    if (i.TimeOfDay.Hours >= 0 && i.TimeOfDay.Hours < 24)
                                    {
                                        intHolidayHrs++;
                                    }
                                }
                            }

                            // reduce Saturday, Sunday and Holidays hours
                            int totMin = Convert.ToInt32(row["TotalMins"].ToString());
                            totMin = totMin - (intHolidayHrs * 60);
 

                            // Verify is less than 4 hrs
                            int intLessFourHrs = (int)(dueDate - DateTime.Now).TotalMinutes;
                            intHolidayHrs = 0;

                            for (var i = dueDate; i < end; i = i.AddHours(1))
                            {
                                if (i.DayOfWeek == DayOfWeek.Saturday | i.DayOfWeek == DayOfWeek.Sunday | (strHolidaysList.Contains("{" + i.Month.ToString("D2") + "-" + i.Day.ToString("D2") + "-" + i.Year + "}")))
                                {
                                    if (i.TimeOfDay.Hours >= 0 && i.TimeOfDay.Hours < 24)
                                    {
                                        intHolidayHrs++;
                                    }
                                }
                            }

                            // reduce Saturday, sunday and holidays hours
                            intLessFourHrs = intLessFourHrs - (intHolidayHrs * 60);

                            // To filter with specific process with timeline for delayed hours
                            if ((row["ActvityName"].ToString() == "Login" || row["ActvityName"].ToString() == "Project Management") && totMin > 60)
                            { row["Current_Activity"] = "Login"; totMin = totMin - 60; }
                            else if ((row["ActvityName"].ToString() == "CleanUp" || row["ActvityName"].ToString() == "Auto-CI") && totMin > 120)
                            { row["Current_Activity"] = "CleanUp"; totMin = totMin - 120; }
                            else if (row["ActvityName"].ToString() == "Pre-Edit" && totMin > 420)
                            { row["Current_Activity"] = "Pre-Edit"; totMin = totMin - 420; }
                            else if ((row["ActvityName"].ToString() == "Normalization" || row["ActvityName"].ToString() == "Auto-ESA" || row["ActvityName"].ToString() == "Mechanical Editing") && totMin > 480)
                            { row["Current_Activity"] = "Normalization"; totMin = totMin - 480; }
                            else if ((row["ActvityName"].ToString() == "Pag" || row["ActvityName"].ToString() == "Auto-Pagination") && intLessFourHrs <= 240)
                            {
                                row["Current_Activity"] = "Pag"; totMin = intLessFourHrs;
                            }
                            else if (row["ActvityName"].ToString() == "ML" && intLessFourHrs <= 240)
                            {
                                row["Current_Activity"] = "ML"; totMin = intLessFourHrs;
                            }
                            else if (row["ActvityName"].ToString() == "CE" && intLessFourHrs <= 240)
                            {
                                row["Current_Activity"] = "CE"; totMin = intLessFourHrs;
                            }
                            else if (row["ActvityName"].ToString() == "TE" && intLessFourHrs <= 240)
                            {
                                row["Current_Activity"] = "TE"; totMin = intLessFourHrs;
                            }
                            else
                            {
                                row["Current_Activity"] = "";
                                //totMin = intLessFourHrs;
                            }

                            //else if (row["Current_Activity"].ToString() == "CE" && row["ESA_Status"].ToString() == "B" && totMin > 1920)
                            //    { row["Current_Activity"] = "CE"; totMin = totMin - 1920;}
                            //else if (row["Current_Activity"].ToString() == "CE" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "" && totMin > 5760)
                            //    { row["Current_Activity"] = "CE"; totMin = totMin - 5760;}
                            //else if (row["Current_Activity"].ToString() == "CE" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "1" && totMin > 8640)
                            //    { row["Current_Activity"] = "CE"; totMin = totMin - 8640;}
                            //else if (row["Current_Activity"].ToString() == "ML" && row["ESA_Status"].ToString() == "A" && totMin > 840)
                            //    { row["Current_Activity"] = "ML"; totMin = totMin - 840;}
                            //else if (row["Current_Activity"].ToString() == "ML" && row["ESA_Status"].ToString() == "B" && totMin > 2280)
                            //    { row["Current_Activity"] = "ML"; totMin = totMin - 2280;}
                            //else if (row["Current_Activity"].ToString() == "ML" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "" && totMin > 6120)
                            //    { row["Current_Activity"] = "ML"; totMin = totMin - 6120;}
                            //else if (row["Current_Activity"].ToString() == "ML" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "1" && totMin > 9000)
                            //    { row["Current_Activity"] = "ML"; totMin = totMin - 9000;}
                            //else if (row["Current_Activity"].ToString() == "Pag" && row["ESA_Status"].ToString() == "A" && totMin > 1200)
                            //    { row["Current_Activity"] = "Pag"; totMin = totMin - 1200;}
                            //else if (row["Current_Activity"].ToString() == "Pag" && row["ESA_Status"].ToString() == "B" && totMin > 2640)
                            //    { row["Current_Activity"] = "Pag"; totMin = totMin - 2640;}
                            //else if (row["Current_Activity"].ToString() == "Pag" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "" && totMin > 6480)
                            //    { row["Current_Activity"] = "Pag"; totMin = totMin - 6480;}
                            //else if (row["Current_Activity"].ToString() == "Pag" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "1" && totMin > 9360)
                            //    { row["Current_Activity"] = "Pag"; totMin = totMin - 9360;}
                            //else if (row["Current_Activity"].ToString() == "TE" && row["ESA_Status"].ToString() == "A" && totMin > 1440)
                            //    { row["Current_Activity"] = "TE"; totMin = totMin - 1440;}
                            //else if (row["Current_Activity"].ToString() == "TE" && row["ESA_Status"].ToString() == "B" && totMin > 2880)
                            //    { row["Current_Activity"] = "TE"; totMin = totMin - 2880;}
                            //else if (row["Current_Activity"].ToString() == "TE" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "" && totMin > 6720)
                            //    { row["Current_Activity"] = "TE"; totMin = totMin - 6720;}
                            //else if (row["Current_Activity"].ToString() == "TE" && row["ESA_Status"].ToString() == "C" && row["CEOnShore"].ToString() == "1" && totMin > 9600)
                            //    { row["Current_Activity"] = "TE"; totMin = totMin - 9600;}
                            //else
                            //     row["Current_Activity"] = "";


                            TimeSpan t = TimeSpan.FromMinutes(totMin);
                            string answer = ((t.Days * 24) + t.Hours) + ":" + string.Format("{0:D2}", (t.Minutes.ToString("D2").Replace("-", "")));
                            if (answer.Contains("-") || !Regex.IsMatch(row["Current_Activity"].ToString(), "^(Pag|CE|ML|TE)$"))
                            {
                                row["Delayed_Hrs"] = "Delay by <span style='font-weight:bold;'>" + answer.Replace("-", "") + "</span> hr/min";

                            }
                            else 
                            {
                                if (DateTime.Now < dueDate)
                                {
                                    row["Delayed_Hrs"] = "Due in " + answer + " hr/min";
                                }
                                else
                                {
                                    row["Delayed_Hrs"] = "Delay by <span style='font-weight:bold;'>" + answer.Replace("-", "") + "</span> hr/min";
                                }
                             
                            }
                            

                            

                        }
                        catch (Exception)
                        {
                        }
                    }


                    DataTable finalTable = (from wp in dtWip.AsEnumerable()
                                          where wp.Field<string>("Current_Activity") != ""
                                          select wp).CopyToDataTable();

                    var tmeJSONString = from a in finalTable.AsEnumerable()
                                     select new[] {
                                         a[0].ToString(),
                                         a[1].ToString(),
                                         a[2].ToString(),
                                         Convert.ToDateTime(a[3].ToString()).ToString("dd-MMM-yyyy HH:mm:ss"),
                                         Convert.ToDateTime(a[4].ToString()).ToString("dd-MMM-yyyy HH:mm:ss"),
                                         a[5].ToString(),
                                         ContentMerge(a[10].ToString().Trim()=="0"?"":a[10].ToString().Trim(), a[13].ToString().Trim()).Trim(),
                                         a[14].ToString().Trim()!=""?Convert.ToDateTime(a[14].ToString()).ToString("dd-MMM-yyyy HH:mm:ss"):a[14].ToString().Trim(),
                                         a[6].ToString(),
                                        "<input type='number' onchange='funRevisedDueDate(this)' value='" + a[7].ToString() + "' id='txt" + a[0] + "_" + a[1] + "_" + a[2] + "' style='width:64px; text-align:center; padding-left: 13px;height:22px;'/>", // "input",//a[7].ToString()
                                        Calc_DueDate_Hrs(a[4].ToString(), a[7].ToString(), strHolidaysList), //a[7].ToString().Trim()!=""?.ToString("dd-MMM-yyyy HH:mm:ss"):a[7].ToString().Trim(),
                                         a[8].ToString()
                                    };
                    

                    List<string> activityOrder = new List<string> { "Login", "CleanUp", "Pre-Edit", "Normalization", "CE", "ML", "Pag","TE", "QC", "Waiting for ZIP File Movement", "Waiting for Smart Track Entry" };   //,"Waiting for ZIP File Movement", "Waiting for Smart Track Entry","Project Management", "CleanUP", "Auto-CI", "Pre-Edit", "Normalization", "Auto-ESA", "Mechanical Editing",
                    List<object[]> activity = new List<object[]>();
                    if (finalTable.Columns.Contains("Current_Activity"))
                    {
                        activity = finalTable.AsEnumerable().GroupBy(g => new { Activity = g["Current_Activity"] }).Select(s => new[] { s.Key.Activity, s.Count() }).OrderByDescending(o => activityOrder.AsEnumerable().Reverse().ToList().IndexOf(o[0].ToString())).ToList();
                    }


                    var activityDic = JsonConvert.SerializeObject(activity);

                    var jsonResult = Json(new { result = "Success", data = tmeJSONString, activity = activityDic }, JsonRequestBehavior.AllowGet);

                    jsonResult.MaxJsonLength = int.MaxValue;

                    return jsonResult;
                    
                }
                catch (Exception)
                {
                    return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult dueDateCalculate(string strCurrDueDate, int intRevHours)
        {
            try
            {
                string strHolidaysList = string.Empty;
                DataSet dsHoliday = new DataSet();
                dsHoliday = DBProc.GetResultasDataSet("SELECT (STUFF((SELECT '{' + CONVERT(varchar,HolidayDate,110) + '}' as 'DATE' FROM tbl_HolidayList FOR XML PATH('')), 1, 0, '')) AS [HoliDays]", Session["sConnSiteDB"].ToString());
                if (dsHoliday.Tables[0].Rows.Count > 0)
                {
                    strHolidaysList = dsHoliday.Tables[0].Rows[0][0].ToString();
                    strHolidaysList = strHolidaysList.Replace("<DATE>", "");
                    strHolidaysList = strHolidaysList.Replace("</DATE>", "");
                }

                DateTime dueDate = DateTime.Parse(strCurrDueDate);

                //dueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day, dueDate.Hour, 0, 0, dueDate.Kind);
                //int intTA = (int)(DateTime.Parse(strCurrDueDate) - dueDate).TotalMinutes;

                for (int i = 1; i <= intRevHours; i++)
                {
                    do
                        dueDate = dueDate.AddHours(1);
                    while (dueDate.DayOfWeek == DayOfWeek.Saturday || dueDate.DayOfWeek == DayOfWeek.Sunday | (strHolidaysList.Contains("{" + dueDate.Month.ToString("D2") + "-" + dueDate.Day.ToString("D2") + "-" + dueDate.Year + "}")));   // | (dueDate.TimeOfDay.Hours >= 0 && dueDate.TimeOfDay.TotalHours <= 7 | dueDate.TimeOfDay.TotalHours >= 21 && dueDate.TimeOfDay.TotalHours <= 24)
                }
                //dueDate = dueDate.AddMinutes(intTA);

                return Json(new { dataResult = dueDate.ToString("dd-MMM-yyyy HH:mm:ss") }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
       
        public string ContentMerge(string strTrack, string strType)
        {
            if (strType == "1" && strTrack != "")
                return strTrack + " - OnShoreCE";
            else
                return strTrack;
        }

        public string Calc_DueDate_Hrs(string dueDate, string DelayHrs, string strHoliday)
        {

            
            DateTime StartDate = Convert.ToDateTime(dueDate);
            if (DelayHrs != "")
            {
                int DelyHrs = Convert.ToInt32(DelayHrs);

                for (int i = 1; i <= DelyHrs; i++)
                {
                    do
                        //StartDate = DateTime.DateAdd(DateInterval.Hour, 1, StartDate);
                        StartDate = StartDate.AddHours(1);
                    while (StartDate.DayOfWeek == DayOfWeek.Saturday | StartDate.DayOfWeek == DayOfWeek.Sunday | (strHoliday.Contains("{" + StartDate.Month.ToString("D2") + "-" + StartDate.Day.ToString("D2") + "-" + StartDate.Year + "}")));
                }
                return StartDate.ToString("dd-MMM-yyyy HH:mm:ss");
            }

            return "";

        }

        public ActionResult FPEstimateEventTrigger(string strCurrActivity, string strJournalAc, string strAutoArtID, string strArticleID, string  RevisedDueDate, int RevisedHrs)
        {
            try
            {
             
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select a.JBM_AutoID,a.AutoArtID,a.ChapterID,s.RevFinStage,s.DueDate from TF_ArticleInfo a INNER JOIN JBM_Info j on a.JBM_AutoID=j.JBM_AutoID join TF_StageInfo s on s.AutoArtID=a.AutoArtID WHERE a.AutoArtID='" + strAutoArtID + "' and s.RevFinStage='FP'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DueDate"].ToString().Contains("00:00:00"))
                    {
                        return Json(new { dataResult = "Incorrect due date in database." }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        DateTime dtExtentDueDate = Convert.ToDateTime(RevisedDueDate);//Convert.ToDateTime(ds.Tables[0].Rows[0]["DueDate"].ToString()).AddHours(RevisedHrs);

                        RevisedHrs = (int)(dtExtentDueDate - DateTime.Parse(ds.Tables[0].Rows[0]["DueDate"].ToString())).TotalHours;

                        string strReturn = DBProc.GetResultasString("Update TF_StageInfo Set DueDate='" + dtExtentDueDate.ToString("dd-MMM-yyyy HH:mm:ss") + "' WHERE RevFinStage='FP' and AutoArtID='" + strAutoArtID + "'; Update TF_ArticleInfo SET ExternalEventActDate=GETDATE(), ExternalEventReplay='" + strCurrActivity + "' WHERE AutoArtID='" + strAutoArtID + "'", Session["sConnSiteDB"].ToString());

                        string strInsertEvent = "INSERT INTO TF_ExternalEventAccess (JBM_AutoID,AutoArtID,ExternalEventID,ExternalEventActDate,Stage,ArticleID,MaxTry,Status,EmpLogin,OperationType) values('" + ds.Tables[0].Rows[0]["JBM_AutoID"].ToString() + "','" + strAutoArtID + "','EV066',GETDATE(),'FP','" + strArticleID + "','0','16','" + Session["UserID"].ToString() + "','" + RevisedHrs + "')";
                        string strResult1 = DBProc.GetResultasString(strInsertEvent, Session["sConnSiteDB"].ToString());

                        string strTrackQuery =  " Insert into " + Init_Tables.gTblProdAccess + "(AutoArtID,EmpAutoID,CustAcc,Descript,AccTime,AccPage,Process) values('" + strAutoArtID + "', '" + Session["EmpAutoID"].ToString() + "', '" + Session["sCustAcc"].ToString() + "', 'Revised TAT and Due Date update', getdate(), 'AM Dashboard', 'Revised TAT Triggered')";
                        strResult1 = DBProc.GetResultasString(strTrackQuery, Session["sConnSiteDB"].ToString());

                    }
                    
                }

                return Json(new { dataResult = "Success" }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { dataResult = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

    }
}