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
using RL = ReferenceLibrary;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Data.OleDb;

namespace SmartTrack.Controllers
{
    [SessionExpire]
    public class SupportController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        DataProc DBProc = new DataProc();
        // GET: Support
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult SignalCreation()
        {
            return View();
        }

        public ActionResult getReports(string articleIds, string filterBy)
        {
            try

            {
                try
                {
                    if (articleIds != "")
                    {
                        articleIds = "" + articleIds.Replace(",", "','").Replace(" ", "','") + "";
                    }

                    if (Session["sCustAcc"].ToString() == "BK" || Session["sCustAcc"].ToString() == "MG")
                    {
                        Init_Tables.gTblChapterOrArticleInfo = "_ChapterInfo";
                    }
                    else { Init_Tables.gTblChapterOrArticleInfo = "_ArticleInfo"; }

                    string strQueryFinal = "Select a.AutoArtID as [Job ID],a.IntrnlID as [Internal ID], d.JBM_ID as [JID],a.JBM_AutoID as AutoId, a.ChapterID as [Article ID], a.iss as [Iss], c.RevFinStage as [Stage],b.PlatformDesc as Platformde, a.DOI,d.SiteID as SiteID,Format (c.LoginRecDate, 'dd MMM yyyy') as [Login Date], Format (c.DueDate, 'dd MMM yyyy') as [Due Date], c.ArtStageTypeID as [Current Stage], '' as [Current Activity], C.ArtStageTypeID,(select CustSN from JBM_CustomerMaster WHERE CustID=d.CustID) as [Cust_SN], '' as [KGLInw], '' as [KGLPro], '' as [KGLProPdf], '' as [WorkingFolder], d.JBM_Intrnl as [JBMIntrnl] from JBM_Platform b inner join " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a on a.PlatformID = b.PlatformID  inner join " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " c on a.AutoArtID=c.AutoArtID inner join Jbm_info d on d.JBM_AutoID=a.JBM_AutoID where   (a.AutoArtid in ('" + articleIds + "') or a.ChapterID in ('" + articleIds + "') or a.IntrnlID in ('" + articleIds + "')) and c.Revfinstage like '" + filterBy + "%' order by LoginEnteredDate Desc";

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    DataTable dtProArtStaDes = DBProc.GetResultasDataTbl("Select DeptActivity,StageDesc,Stage from " + Init_Tables.gTblProdArtStatusDesDept + "", Session["sConnSiteDB"].ToString());
                    DataTable dtEmpDet = DBProc.GetResultasDataTbl("Select EmpLogin,EmpAutoID,EmpName,EmpSurname from " + Init_Tables.gTblEmployee + " ", Session["sConnSiteDB"].ToString());

                    string CurrStage = string.Empty;
                    int errorRow = 0;
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        ReportsController wObj = new ReportsController();
                        for (int i = 0, loopTo1 = ds.Tables[0].Rows.Count - 1; i <= loopTo1; i++)
                        {
                            errorRow = i;
                            string tmpFromActivity = "";
                            string tmpCurActivity = "";
                            string tmpToActivity = "";

                            if (ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString() != "")
                            {
                                CurrStage = "";
                                if (ds.Tables[0].Columns.Contains("Current Stage"))
                                {
                                    CurrStage = wObj.fn_CurrentstageFull(ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString(), dtProArtStaDes, dtEmpDet);
                                    ds.Tables[0].Rows[i]["Current Stage"] = CurrStage;
                                }

                                if (ds.Tables[0].Columns.Contains("Current Activity"))
                                {
                                    if (ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(0, 2) == "_S")
                                    {
                                        ds.Tables[0].Rows[i]["Current Activity"] = wObj.CurrentActivity(dtProArtStaDes.Select("DeptActivity = '" + ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(7, 5) + "' and DeptActivity is not Null")[0][1].ToString(), ds.Tables[0].Rows[i]["pagcorr"].ToString(), ds.Tables[0].Rows[i]["TeCorr"].ToString());
                                    }
                                    else if (CurrStage.Contains("_Hold")) // ' Code added on 22-Aug-2018 as per Mr. Siva instruction
                                    {
                                        ds.Tables[0].Rows[i]["Current Activity"] = "Hold";
                                        tmpFromActivity = dtProArtStaDes.Select("DeptActivity = '" + ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(0, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                        tmpCurActivity = "Hold";
                                    }
                                    else
                                    {
                                        string strCurActivity = "";
                                        string strFromActivity = "";
                                        string strToActivity = "";
                                        try
                                        {
                                            strFromActivity = dtProArtStaDes.Select("DeptActivity = '" + ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(0, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                            strToActivity = dtProArtStaDes.Select("DeptActivity = '" + ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(12, 5) + "' and DeptActivity is not Null")[0][1].ToString();
                                            strFromActivity = IIf((strFromActivity ?? "") == (strToActivity ?? ""), "", strFromActivity).ToString();
                                            tmpFromActivity = strFromActivity;
                                            tmpToActivity = strToActivity;
                                            strFromActivity = IIf(string.IsNullOrEmpty(strFromActivity), "", " (" + strFromActivity + ")").ToString();
                                        }
                                        catch (Exception ex)
                                        {
                                            strFromActivity = "";
                                        }

                                        strCurActivity = wObj.CurrentActivity(dtProArtStaDes.Select("DeptActivity = '" + ds.Tables[0].Rows[i]["ArtstageTypeID"].ToString().Substring(12, 5) + "' and DeptActivity is not Null")[0][1].ToString(), "", "");
                                        if (strCurActivity.ToLower().Contains("corr"))
                                        {
                                            strCurActivity = strCurActivity + strFromActivity;
                                        }

                                        tmpCurActivity = strCurActivity;

                                        ds.Tables[0].Rows[i]["Current Activity"] = strCurActivity;

                                    }



                                }
                            }
                            else
                            {
                                if (ds.Tables[0].Columns.Contains("Current Stage"))
                                {
                                    ds.Tables[0].Rows[errorRow]["Current Stage"] = "";
                                }

                                if (ds.Tables[0].Columns.Contains("Current Activity"))
                                {
                                    ds.Tables[0].Rows[errorRow]["Current Activity"] = "";
                                }
                            }

                            // To assign path in NAS icon
                            RL.ArtDet A = new RL.ArtDet();

                            string strVolIssNo = "";
                            string strVolDir = string.Empty;
                            string strAllocVolIss = "Vol00000";

                            if (ds.Tables[0].Rows[i]["Iss"].ToString() != null || ds.Tables[0].Rows[i]["Iss"].ToString() != "")
                            {
                                strVolIssNo = ds.Tables[0].Rows[i]["Iss"].ToString();
                                if (strVolIssNo != "")
                                {
                                    strVolDir = RL.clsFileIO.Proc_Extract_VolDir(strVolIssNo);
                                }
                            }
                            else
                                strVolDir = strAllocVolIss;

                            A.AD.VolDir = strVolDir;
                            A.AD.JrnlSiteID = ds.Tables[0].Rows[i]["SiteId"].ToString();
                            A.AD.InternalID = ds.Tables[0].Rows[i]["Internal ID"].ToString();
                            A.AD.InternalJID = ds.Tables[0].Rows[i]["JID"].ToString();
                            A.AD.CustSN = ds.Tables[0].Rows[i]["Cust_SN"].ToString();
                            A.AD.CustAccess = "TF";
                            A.AD.CustGroup = "";
                            A.AD.ConnString = DBProc.DbConn_Assign2RL(Session["sConnSiteDB"].ToString());
                            String strCeStylesheetDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F50-01", false);


                        }
                    }



                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {"",
                                         a[0].ToString(),
                                         a[13].ToString(),
                                         a[2].ToString(),
                                         a[3].ToString(),
                                         a[4].ToString(),
                                         a[5].ToString(),
                                         a[6].ToString(),
                                         a[7].ToString(),
                                         a[8].ToString(),
                                         a[9].ToString(),
                                         a[10].ToString(),
                                         a[11].ToString(),
                                         a[12].ToString(),
                                         a[1].ToString(),
                                         a[14].ToString(),
                                         a[15].ToString(),
                                         a[16].ToString(),
                                         a[17].ToString(),
                                         a[18].ToString(),
                                         a[19].ToString()

                                    };
                    return Json(new { dataResult = JSONString }, JsonRequestBehavior.AllowGet);
                }
                catch (Exception ex)
                {
                    return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult generateSignals(string jObject)
        {
            try
            {
                var ArtObject = JsonConvert.DeserializeObject<ArtObject>(jObject);


                //List<string> sigList = JsonConvert.DeserializeObject<List<string>>(jObject);
                //if (sigList.Count > 0)
                //{
                //    for (int i = 0; i < sigList.Count; i++)
                //    {
                //        string sArticleID = sigList[i].Split('|')[0];
                //    }
                //}

                return Json(new { dataResult = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataResult = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        #region ActivityChange 

        public ActionResult ActivityChange()
        {
            string strQuery = "select CustID,CustSN from JBM_CustomerMaster where CustType='" + Session["sCustAcc"].ToString() + "' order by CustSN";
            DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
            ViewBag.CustList = from a in dt.AsEnumerable() select new SelectListItem { Text = a["CustSN"].ToString(), Value = a["CustID"].ToString() };

            strQuery = "select WFName from JBM_WFCode order by Convert(int,Replace(WFName,'W',''))"; // where is_CustStage is null
            dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
            ViewBag.WF = from a in dt.AsEnumerable() select new SelectListItem { Text = a["WFName"].ToString(), Value = a["WFName"].ToString() };

            return View();
        }
        public ActionResult getJournalID(string custID)
        {
            try
            {
                string strCustID = "";
                if (custID != "-1")
                {
                    strCustID = " and c.CustID='" + custID + "'";
                }

                string strQuery = "select JBM_ID, JBM_AutoID from JBM_Info i inner join JBM_CustomerMaster c on c.CustID = i.CustID where c.CustType = '" + Session["sCustAcc"].ToString() + "' and i.JBM_Disabled = '0' " + strCustID + " order by JBM_ID";

                DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                var result = Newtonsoft.Json.JsonConvert.SerializeObject(dt);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("");
            }

        }

        public ActionResult getStageDetails(string custID, string JBM_AutoID, string articleID)
        {
            try
            {

                string articleIDList = "";

                if (articleID.Trim() != "")
                {

                    articleID = Regex.Replace(articleID, @"(\n|,|\s|\;)", "|", RegexOptions.IgnoreCase);
                    articleID = articleID.Replace("||", "|");
                    articleIDList = ("'" + articleID.Trim().Replace("|", "','") + "'");
                }

                if (articleIDList != "")
                {
                    articleIDList = " and (a.AutoArtID in (" + articleIDList + ") or a.ChapterID in (" + articleIDList + ") or a.IntrnlID in (" + articleIDList + "))";
                }

                string strColNamePrj = string.Empty;
                string strColNameChap = string.Empty;
                if (Session["sCustAcc"].ToString() == "BK" || Session["sCustAcc"].ToString() == "MG")
                {
                    Init_Tables.gTblChapterOrArticleInfo = "_ChapterInfo";
                    strColNamePrj = "Project ID";
                    strColNameChap = "Chapter ID";
                }
                else
                {
                    Init_Tables.gTblChapterOrArticleInfo = "_ArticleInfo";
                    strColNamePrj = "Journal ID";
                    strColNameChap = "Job ID";
                }

                string strQuery = "select '' as tblCheckBox,i.JBM_id as [" + strColNamePrj + "],a.ChapterID as [" + strColNameChap + "],a.AutoArtID,RevFinStage as Stage,dbo.fn_Currentstage(s.ArtstageTypeID) as [Current Process],s.Rev_Wf as [Work Flow],s.ArtstageTypeID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join " + Init_Tables.gTblJrnlInfo + " i on i.JBM_AutoID=a.JBM_AutoID inner join JBM_CustomerMaster c on c.CustID=i.CustID inner join " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " s on s.AutoArtID=a.AutoArtID where c.CustID = ISNULL(NULLIF(convert(varchar(10),'" + custID + "'),''),c.CustID) and i.JBM_AutoID = ISNULL(NULLIF(convert(varchar(10),'" + JBM_AutoID + "'),''),i.JBM_AutoID)  and s.DispatchDate is null and a.WIP=1 " + articleIDList;

                DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                DataView dv = new DataView(dt);
                //DataTable dtWF = dv.ToTable(true, "Work Flow");

                var result = Newtonsoft.Json.JsonConvert.SerializeObject(dt);
                var resultWF = (from r in dt.AsEnumerable() select r["Work Flow"]).Distinct().ToList();

                var jsonResult = Json(new { data = result, WF = resultWF }, JsonRequestBehavior.AllowGet);

                jsonResult.MaxJsonLength = int.MaxValue;

                return jsonResult;
            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

        public ActionResult getProcessList(string strWF)
        {
            try
            {
                //string strQuery = "if exists(select * from sys.all_columns where object_id in (select object_id from sys.tables where name='JBM_Info') and name='"+ stageID + "_wf') begin declare @WFCode varchar(200)='' (select top 1 @WFCode =REPLACE(WFCode,'|',',') from JBM_Info i inner join JBM_WFCode w on w.WFName=i."+ stageID + "_wf where JBM_AutoID='"+ JBM_AutoID + "') declare @temp table ( WF int ) insert into @temp SELECT distinct Split.A.value('.', 'NVARCHAR(MAX)') [DATA] FROM ( SELECT CAST('<X>'+REPLACE(@WFCode, ',', '</X><X>')+'</X>' AS XML) AS String ) AS A CROSS APPLY String.nodes('/X') AS Split(A); select DeptActivity,StageDesc from @temp t inner join JBM_ProdArtStatDesDept p on p.DeptActivity= FORMAT(WF,'S1000') where WF !=0 order by DeptActivity end else select DeptActivity,StageDesc from JBM_ProdArtStatDesDept where DeptActivity is null ";

                string strQuery = "declare @WFCode varchar(200)='' declare @temp table ( WF int ) select @WFCode =REPLACE(WFCode,'|',',') from JBM_WFCode where WFName='" + strWF + "'  insert into @temp SELECT distinct Split.A.value('.', 'NVARCHAR(MAX)') [DATA] FROM ( SELECT CAST('<X>'+REPLACE(@WFCode, ',', '</X><X>')+'</X>' AS XML) AS String ) AS A CROSS APPLY String.nodes('/X') AS Split(A); select DeptActivity,StageDesc from @temp t inner join JBM_ProdArtStatDesDept p on p.DeptActivity= FORMAT(WF,'S1000') where WF !=0 order by DeptActivity ";

                DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                var result = Newtonsoft.Json.JsonConvert.SerializeObject(dt);
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("");
            }

        }

        public string updateActivity(string autoArtIDlist, string processID, string WF, string trackingDetails, List<string[]> lst, string strProcessID)
        {
            try
            {
                string ArtStageTypeID = "'" + processID + "_S1009_" + processID + "_00001_" + "'+convert(varchar(100),getdate(),121)";

                string deleteAllocationQry = "";
                string trackingQry = "";
                string strQuery = "";

                foreach (string[] arr in lst)
                {
                    string autoArtID = arr[0].ToString();
                    string stage = arr[1].ToString();
                    string currentWF = arr[2].ToString();
                    string currentProcess = arr[3].ToString();
                    string currentArtStageTypeID = arr[4].ToString();
                    string currentStageR = currentArtStageTypeID.Substring(12, 5);
                    string trackingDesc = "";
                    string updateFields = "";

                    if (WF != currentWF)
                    {
                        trackingDesc = "Workflow updated from " + currentWF + " to " + WF + "";
                    }

                    if (processID != "")
                    {
                        trackingDesc = trackingDesc != "" ? trackingDesc + " and " : trackingDesc;
                        trackingDesc = trackingDesc + " Article wise stage updated from " + currentProcess + " to " + strProcessID + "";

                        //ArtStageTypeID = " CurrentStatus=null,CurrentProcess=null, ArtStageTypeID='" + processID + "_S1009_" + processID + "_00001_" + "'+convert(varchar(100),getdate(),121)";

                        if (processID == "S1016" | processID == "S1020")
                        {
                            updateFields = " ,GraphArtStageTypeID=" + ArtStageTypeID + "";
                        }
                        else
                        {
                            updateFields = " ,ArtStageTypeID=" + ArtStageTypeID + "";
                        }

                        if (processID == "S1003")
                        {
                            updateFields = updateFields + ",CE_Mode=null,Copyeditor=null";
                        }
                    }

                    if (trackingDesc != "")
                    {
                        //trackingQry = trackingQry != "" ? trackingQry + "," : trackingQry;
                        //trackingQry = trackingQry + "('"+autoArtID+"','"+Session["EmpAutoID"].ToString()+ "','" + Session["custAcc"].ToString() + "','"+ trackingDesc + "',getdate(),'Activity/WF Change','Activity/WF Change'))";

                        //updateFields = updateFields + ",CeArtStageTypeID=case when ISNULL(CeArtStageTypeID,'')<>'' then " + ArtStageTypeID + " else CeArtStageTypeID end, XMLArtStageTypeID = case when ISNULL(XMLArtStageTypeID,'') <> '' then " + ArtStageTypeID + " else XMLArtStageTypeID end ";

                        strQuery = strQuery + " Update " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " set Rev_Wf='" + WF + "' " + updateFields + ",CeArtStageTypeID=case when ISNULL(CeArtStageTypeID,'')<>'' then " + ArtStageTypeID + " else CeArtStageTypeID end, XMLArtStageTypeID = case when ISNULL(XMLArtStageTypeID,'') <> '' then " + ArtStageTypeID + " else XMLArtStageTypeID end " + " where autoartid='" + autoArtID + "' and RevFinStage='" + stage + "' and DispatchDate is null; ";

                        strQuery = strQuery + " update " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " set ArtStageTypeID=" + ArtStageTypeID + " where AutoArtID='" + autoArtID + "' and Stage='" + stage + "'; ";

                        if (currentStageR.Contains("S1017") | Regex.IsMatch(currentArtStageTypeID, "(S1026|S1050|S1051|S1052)"))
                        {
                            strQuery = strQuery + " delete from " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " where autoartid='" + autoArtID + "' and stage='" + stage + "' and deptCode not in('20','90','100'); ";
                        }

                        strQuery = strQuery + " insert into " + Init_Tables.gTblProdAccess + "(AutoArtID,EmpAutoID,CustAcc,Descript,AccTime,AccPage,Process) values('" + autoArtID + "', '" + Session["EmpAutoID"].ToString() + "', '" + Session["sCustAcc"].ToString() + "', '" + trackingDesc + "', getdate(), 'Activity/WF Change', 'Activity/WF Change');";
                    }

                    //updateFields = updateFields + ",CeArtStageTypeID=case when ISNULL(CeArtStageTypeID,'')<>'' then " + ArtStageTypeID + " else CeArtStageTypeID end, XMLArtStageTypeID = case when ISNULL(XMLArtStageTypeID,'') <> '' then " + ArtStageTypeID + " else XMLArtStageTypeID end ";
                }

                //if(trackingQry != "")
                //{
                //    trackingQry = "insert into " + Init_Tables.gTblProdAccess + " (AutoArtID,EmpAutoID,CustAcc,Descript,AccTime,AccPage,Process) values" + trackingQry + "";
                //} 

                //string strQuery = "Update " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " set Rev_Wf='" + WF + "' " + updateFields + " where " + autoArtIDlist + "  and DispatchDate is null " + trackingQry + deleteAllocationQry + "";

                string result = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                //int result = 0;

                if (result != "")
                {
                    return "Success";
                }
                else
                {
                    return "Failed";
                }
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        public ActionResult ManualIngestion()
        {
            return View();
        }

        public ActionResult ManualIngestion_FileUpload(string strCustSN, List<HttpPostedFileBase> strFiles)
        {
            try
            {
                RL.ArtDet aDet = new RL.ArtDet();
                aDet.AD.CustSN = strCustSN;
                aDet.AD.CustAccess = Session["sCustAcc"].ToString();
                aDet.AD.ConnString = DBProc.getConnection(Session["sConnSiteDB"].ToString()).ConnectionString;

                string strRootPath = RL.clsFileIO.Proc_Get_Directory_Path(ref aDet, "F5000-Incoming", false);
                string strRootLogPath = RL.clsFileIO.Proc_Get_Directory_Path(ref aDet, "F5000-Logs", false);
                string strError = "";

                if (strRootPath == "")
                {
                    return Json(new { result = "Error", data = "File upload path is missing" }, JsonRequestBehavior.AllowGet);
                }

                strRootPath = Path.Combine(strRootPath, "DOI_XML_Signals", "In");
                strRootLogPath = Path.Combine(strRootLogPath, "ManualIngestion", "Log.txt");

                foreach (HttpPostedFileBase strFile in strFiles)
                {
                    string strFullName = Path.Combine(strRootPath, strFile.FileName);

                    if (System.IO.File.Exists(strFullName))
                    {
                        strError = strError + "," + strFile.FileName;
                    }
                    else
                    {
                        strFile.SaveAs(strFullName);

                        System.Text.StringBuilder sb = new System.Text.StringBuilder();
                        sb.Append(",{");
                        sb.Append("\"JournalName\":\"\",");
                        sb.Append("\"FileType\":\"ManualIngestion_DOIXML\",");
                        sb.Append("\"FileName\":\"" + strFile.FileName + "\",");
                        sb.Append("\"ReceivedDate\":\"" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "\",");
                        sb.Append("\"FilePath\":\"" + strRootPath.Replace(@"\", @"\\") + "\",");
                        sb.Append("\"EmpAutoID\":\"" + Session["EmpAutoId"].ToString() + "\",");
                        sb.Append("\"EmpName\":\"" + Session["EmpName"].ToString() + "\",");
                        sb.Append("\"CustAccess\":\"" + Session["sCustAcc"].ToString() + "\",");
                        sb.Append("\"SiteID\":\"" + Session["sSiteID"].ToString() + "\"");
                        sb.Append("}");

                        System.IO.File.AppendAllText(strRootLogPath, sb.ToString());
                    }
                }

                return Json(new { result = "Success", data = strError.TrimStart(',') }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult ManualIngestion_Details(string CustSN)
        {
            try
            {
                RL.ArtDet aDet = new RL.ArtDet();
                aDet.AD.CustSN = CustSN;
                aDet.AD.CustAccess = Session["sCustAcc"].ToString();
                aDet.AD.ConnString = DBProc.getConnection(Session["sConnSiteDB"].ToString()).ConnectionString;

                string strRootLogPath = RL.clsFileIO.Proc_Get_Directory_Path(ref aDet, "F5000-Logs", false);
                strRootLogPath = Path.Combine(strRootLogPath, "ManualIngestion", "Log.txt");
                string strLogs = "";

                if (System.IO.File.Exists(strRootLogPath))
                {
                    strLogs = System.IO.File.ReadAllText(strRootLogPath);
                }

                if (strLogs.Trim() == "")
                {
                    return Json(new { result = "Success", data = "" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    strLogs = "[" + strLogs.Trim().TrimStart(',') + "]";

                    DataTable table = JsonConvert.DeserializeObject<DataTable>(strLogs);

                    string showAll = Session["gJwAccItm"].ToString().Contains("|IMA|") ? "" : " AND EmpAutoID='" + Session["EmpAutoID"].ToString() + "'";

                    DataView dv = new DataView(table, @"CustAccess = '" + Session["sCustAcc"].ToString() + "' AND SiteID = '" + Session["sSiteID"].ToString() + "'" + showAll, "ReceivedDate desc", DataViewRowState.CurrentRows);

                    var jsonData = JsonConvert.SerializeObject(dv.ToTable());
                    var jsonResult = Json(new { result = "Success", data = jsonData }, JsonRequestBehavior.AllowGet);
                    jsonResult.MaxJsonLength = int.MaxValue;
                    return jsonResult;
                }
            }
            catch (Exception ex)
            {
                return Json(new { result = "Error", data = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        #endregion


        object IIf(bool expression, object truePart, object falsePart)
        {
            return expression ? truePart : falsePart;
        }

    }

    public class ArtObject
    {
        public string vAutoArtID { get; set; }
        public string vJID { get; set; }
        public string vJBMIntrnl { get; set; }
        public string vArticleID { get; set; }
        public string vJBMAutoID { get; set; }
        public string vIss { get; set; }
        public string vCurStage { get; set; }
        public string vCurrActivity { get; set; }
        public string vInternalID { get; set; }
        public string vLoginDate { get; set; }
        public string vDueDate { get; set; }
        public string vSiteID { get; set; }
        public string vCustSN { get; set; }
        public string vCopyDoc { get; set; }
    }
}