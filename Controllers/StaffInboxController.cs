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

namespace SmartTrack.Controllers
{
    [SessionExpire]
    public class StaffInboxController : Controller
    {
        clsWFProcedure clsWFProc = new clsWFProcedure();
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        DataTable dts = new DataTable();
        [SessionExpire]
        // GET: StaffInbox
        public ActionResult Index()  //?CustAcc=BK&EmpID=40385&SiteID=L0003&DeptCode=30
        {
          
            if (Session["sCustAcc"].ToString() == "BK" || Session["sCustAcc"].ToString() == "MG")
            {
                Init_Tables.gTblChapterOrArticleInfo = "_ChapterInfo";
            }
            else { Init_Tables.gTblChapterOrArticleInfo = "_ArticleInfo"; }


            //To load the department list in dropdown.
            string strDeptAcc = Session["DeptAcc"].ToString();
            string sDeptAcc = "";
            sDeptAcc = strDeptAcc.Substring(1, strDeptAcc.Length - 2).ToString().Replace("|", ",");

            DataSet ds = new DataSet();

            List<SelectListItem> collcDept = new List<SelectListItem>();
            ds = DBProc.GetResultasDataSet("Select DeptCode,DeptName,DeptSN from JBM_DepartmentMaster where DeptCode in (" + sDeptAcc + ")", Session["sConnSiteDB"].ToString());

            for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
            {
                string strDeptCode = ds.Tables[0].Rows[intCount]["DeptCode"].ToString();
                string strDeptName = ds.Tables[0].Rows[intCount]["DeptName"].ToString();

                collcDept.Add(new SelectListItem
                {
                    Text = strDeptName,
                    Value = strDeptCode
                });
            }
            ViewBag.lstDept = collcDept;
            return View();

        }
        //[ChildActionOnly]
        //public ActionResult RenderMenu()
        //{
        //    Session["sIss"] = "sk";
        //    Session["sStage"] = "1.7";
        //    return PartialView("_PartialArticleDetails");
        //}
        public ActionResult GetPickedArticleDetails(string artDetails)
        {
            try
            {
                //ReferenceLibrary.ArtDet A = new ReferenceLibrary.ArtDet();
                
                //A.AD.InwardDirPath = ReferenceLibrary.clsFileIO.Proc_Get_Directory_Path(A, "", false);
                

                List<string> lstArtDetails = JsonConvert.DeserializeObject<List<string>>(artDetails);
                if (lstArtDetails.Count > 0)
                {
                    for (int i = 0; i < lstArtDetails.Count; i++)
                    {

                        string sJAutoID = ""; string sJrnlName = ""; string sArticleID = ""; string sStage = ""; string sIss = ""; string sDueDate = ""; string sWF = "";
                        sJAutoID = lstArtDetails[i].Split('|')[0].ToString();
                        sJrnlName = lstArtDetails[i].Split('|')[1].ToString();
                        sArticleID = lstArtDetails[i].Split('|')[2].ToString();
                        sStage = lstArtDetails[i].Split('|')[3].ToString();
                        sIss = lstArtDetails[i].Split('|')[4].ToString();
                        sDueDate = lstArtDetails[i].Split('|')[5].ToString();
                        sWF = lstArtDetails[i].Split('|')[6].ToString();

                        // Create the article details session
                        //Session["sJAutoID"] = sJAutoID;
                        //Session["sJrnlName"] = sJrnlName;
                        //Session["sArticleID"] = sArticleID;
                        //Session["sIss"] = sIss;
                        //Session["sStage"] = sStage;
                        //Session["sDueDate"] = sDueDate;
                        //Session["sWF"] = sWF;

                        string strInstructionHtml = string.Empty;
                        strInstructionHtml = "<tr class='text-primary' align='left' valign='top' style='font-weight:normal;'><th align='left' scope='col' style='width:150px;'>Date</th><th align='left' scope='col'>Instruction</th></tr>";
                        DataSet ds = new DataSet();

                        //List<SelectListItem> collcInstr = new List<SelectListItem>();
                        ds = DBProc.GetResultasDataSet("Select CONVERT(varchar,InstDate,100) as [InstructionDate],Instruction from [dbo].[" + Session["sCustAcc"].ToString() +  "_SplInstructions] WHERE AutoArtID='" + sJAutoID + "' and Stage='" + sStage + "'", Session["sConnSiteDB"].ToString());

                        //for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                        //{
                        //    string strDate = ds.Tables[0].Rows[intCount]["InstDate"].ToString();
                        //    string strInstr = ds.Tables[0].Rows[intCount]["Instruction"].ToString();

                        //    strInstructionHtml += "<tr><td style='width:150px;'>" + strDate  + "</td><td>" + strDate + "</td></tr>";
                        //    //collcInstr.Add(new SelectListItem
                        //    //{
                        //    //    Text = strInstr,
                        //    //    Value = strDate
                        //    //});
                        //}
                        //ViewBag.lstSplInstr = collcInstr;

                        string strArtDetails= "<tbody><tr><td class='header'>Job Auto ID:&nbsp;<span id='lblJAutoID' style='color:Black;'>" + sJAutoID + "</span></td><td class='header'>Journal ID:&nbsp;<span id = 'lblJrnlName' style='color:Black;'>" + sJrnlName + "</span></td><td class='header'>Article ID:&nbsp;<span id = 'lblArticleID' style='color:Black;'>" + sArticleID + "</span></td><td class='header'>Iss:&nbsp;<span id = 'lblIss' style='color:Black;'>" + sIss + "</span></td><td class='header'>Stage:&nbsp;<span id = 'lblStage' style='color:Black;'>" + sStage + "</span></td><td class='header'>Due Date:&nbsp;<span id = 'lblDuedate' style='color:Black;'>" + sDueDate + "</span></td><td class='header'>WF:&nbsp;<span id = 'lbl_WF' style='color:Black;'>" + sWF + "</span></td></tr></tbody>";


                        var JSONString = from a in ds.Tables[0].AsEnumerable()
                                         select new[] {a[0].ToString(),a[1].ToString()};

                        return Json(new { dataDB = strArtDetails, dataInstr = JSONString }, JsonRequestBehavior.AllowGet);

                       // return Content("<tbody><tr><td class='header'>Job Auto ID:&nbsp;<span id='lblJAutoID' style='color:Black;'>" + sJAutoID + "</span></td><td class='header'>Journal ID:&nbsp;<span id = 'lblJrnlName' style='color:Black;'>" + sJrnlName + "</span></td><td class='header'>Article ID:&nbsp;<span id = 'lblArticleID' style='color:Black;'>" + sArticleID + "</span></td><td class='header'>Iss:&nbsp;<span id = 'lblIss' style='color:Black;'>" + sIss + "</span></td><td class='header'>Stage:&nbsp;<span id = 'lblStage' style='color:Black;'>" + sStage + "</span></td><td class='header'>Due Date:&nbsp;<span id = 'lblDuedate' style='color:Black;'>" + sDueDate + "</span></td><td class='header'>WF:&nbsp;<span id = 'lbl_WF' style='color:Black;'>" + sWF + "</span></td></tr></tbody>", "text/html");

                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

               

                return Json(new { dataDB = "Success" }, JsonRequestBehavior.AllowGet);
          


            }
            catch (Exception)
            {

                return Json(new { dataDB = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult FileExplorerList(string fileExplorerPath)
        {
            fileExplorerPath = fileExplorerPath.Replace("[$]", @"\").Replace("[@@@]", "#");

            if (fileExplorerPath == "")
            {
                return Json(new { dataDB = "Path does not exist." }, JsonRequestBehavior.AllowGet);
            }

            DataTable dtbl = new DataTable();
            dtbl = DBProc.GetDataTableFileExplorer(fileExplorerPath);

            DataSet ds = new DataSet();
            ds.Tables.Add(dtbl);

            var JSONString = from a in ds.Tables[0].AsEnumerable()
                             select new[] {a[0].ToString(), a[1].ToString(), a[2].ToString(), a[3].ToString(), a[4].ToString(), a[5].ToString(), a[6].ToString()
                                         };

            return Json(new { dataDB = JSONString }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetInboxAlloatedDetails()
        {
            try
            {
                string strFilterResult = "";
                string strDispFilterQry = "";
                string strFailureFilter = "";
                string strMEFilterResult = "";
                string StrOnshoreMeFilterRes = "";

                 strDispFilterQry = " and s.DispatchDate is null ";
                string strArtLocFilter = "";

                if (Session["sCustAcc"].ToString() == "TF" & (Request.Url.ToString().ToLower().Contains("10.18") | Request.Url.ToString().ToLower().Contains("cps.kwglobal.com") | Request.Url.ToString().ToLower().Contains("smarttrack.cenveo.com/smarttrack-ch")))
                    strArtLocFilter = " and (j.SiteID='" + Init_SiteID.gSiteChennai + "' OR RIGHT(s.ArtPickup_Loc, 5)= '" + Init_SiteID.gSiteChennai + "' ) ";
                else if (Session["sCustAcc"].ToString() == "TF" & (Request.Url.ToString().Contains("10.21.") | Request.Url.ToString().ToLower().Contains("cpsblr.cenveo.com") | Request.Url.ToString().ToLower().Contains("localhost")))
                    strArtLocFilter = " and (j.SiteID='" + Init_SiteID.gSiteBangalore + "' OR RIGHT(s.ArtPickup_Loc, 5)= '" + Init_SiteID.gSiteBangalore + "' ) ";

                switch (Session["DeptCode"].ToString())
                {
                    case  Init_Department.gDcImInward:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1070%') " + "or (s.ArtStageTypeID like '%_S1070%'))";
                            break;
                        }

                    case Init_Department.gDcInward:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1001%') " + "or (s.ArtStageTypeID like '%_S1001%') or (s.ArtStageTypeID like '%_S1070%'))";
                            break;
                        }

                    case Init_Department.gDcIm:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1027%') or (s.ArtStageTypeID like '%_S1027%') or (s.ArtStageTypeID like '%_S1029%')) and (SUBSTRING(s.ArtStageTypeID,1,35) not like '%S1057_S1027%')";
                           // strDispFilterQry = "";
                            break;
                        }

                    case Init_Department.gDcConversion:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1113%') " + "or (s.ArtStageTypeID like '%_S1113%'))";
                            break;
                        }

                    case Init_Department.gDcTechSupport:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1110%') " + "or (s.ArtStageTypeID like '%_S1110%'))";
                            break;
                        }

                    case Init_Department.gDcRE:
                        {
                            // strFilterResult = "and ( (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1002%') " &
                            // "or (e.ArtStageTypeID like '%_S1002%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1152%') " &
                            // "or (e.ArtStageTypeID like '%_S1152%') )"
                            strFilterResult = "and ( (SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1009%') " + "or (s.ArtStageTypeID like '%_S1009%'))";
                            //strFailureFilter = "or ((e.ArtStageTypeID like '%_S1161_S1002%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1002%'))) "; // &
                            break;
                        }

                    case Init_Department.gDcCE:
                        {
                            if (Session["DeptCode"].ToString() == "TF")
                            //    strMEFilterResult = " OR (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1152%') " + "or (e.ArtStageTypeID like '%_S1152%') AND b.CurrentProcess='S1152' ";

                            //StrOnshoreMeFilterRes = " OR (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1187%') OR (e.ArtStageTypeID like '%_S1187_%') AND b.CurrentProcess='S1187'";

                            
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1003%') " + "or (s.ArtStageTypeID like '%_S1003%') Or (SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1057_S1057_%') or (SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1038%') " + "or (s.ArtStageTypeID like '%_S1038%') ) ";
                            //strFailureFilter = " or (e.ArtStageTypeID like '%_S1161_S1003%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1003%')) ";
                            break;
                        }

                    case Init_Department.gDcXML:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1004%') " + "or (s.ArtStageTypeID like '%_S1004%'))";
                         //   strFailureFilter = " or (e.ArtStageTypeID like '%_S1161_S1004%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1004%')) ";
                            break;
                        }

                    case  Init_Department.gDcMyPet:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1041%') " + "or (s.ArtStageTypeID like '%_S1041%'))";
                            break;
                        }

                    case Init_Department.gDcICE:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1079%') " + "or (s.ArtStageTypeID like '%_S1079%'))";
                            break;
                        }

                    case Init_Department.gDcPag:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1005%') " + "or (s.ArtStageTypeID like '%_S1005%'))";
                         //   strFailureFilter = " or (e.ArtStageTypeID like '%_S1161_S1005%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1005%')) ";
                            break;
                        }

                    case Init_Department.gDcTE:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1006%') " + "or (s.ArtStageTypeID like '%_S1006%'))";
                            //strFailureFilter = " or (e.ArtStageTypeID like '%_S1161_S1006%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1006%')) ";
                            break;
                        }

                    case Init_Department.gDcQC:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1007%') " + "or (s.ArtStageTypeID like '%_S1007%'))";
                            //strFailureFilter = " or (e.ArtStageTypeID like '%_S1161_S1007%') or (SUBSTRING(b.ArtStageTypeID,1,35) like '%_S1161_S1007%')) ";
                            break;
                        }

                    case Init_Department.gDcTemplate:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1024%') " + "or (s.ArtStageTypeID like '%_S1024%'))";
                            break;
                        }

                    case Init_Department.gDcEphrata:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1028%') " + "or (s.ArtStageTypeID like '%_S1028%') or (s.ArtStageTypeID like '%_S1030%'))";
                            break;
                        }

                    case Init_Department.gDcPPC:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1119%') " + "or (s.ArtStageTypeID like '%_S1119%'))";
                            break;
                        }

                    case Init_Department.gDcProjectManagement:
                    case Init_Department.gDcAM:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1045%') " + "or (s.ArtStageTypeID like '%_S1045%'))";
                            break;
                        }

                    case Init_Department.gDcAltText:
                        {
                            strFilterResult = "and ((SUBSTRING(s.ArtStageTypeID,1,35) like '%_S1128%') " + "or (s.ArtStageTypeID like '%_S1128%'))";
                            break;
                        }

                    case  Init_Department.gDcGraphics:
                        {
                            break;
                        }
                }


                //string strQuery = "Select  AllocatedTo, AssignedTo,ManuscriptID,  AutoArtID, AllocatedDate, Pages, Status,DeptCode, VolIss,  InternalRemarks, Stage, AllocatedBy, AllocatedByName, SplitNo, PlatformDesc, JBM_Intrnl, CustSN, [CurrentStage], ArtStageTypeID, GraphArtStageTypeID, CeArtStageTypeID, XMLArtStageTypeID, CeParallel, XMLParallel, NumofFigures, IntrnlID, WfCode, CurrentProcess, InitialCE, JBM_AutoID, MyPet, DueDate, TeCorr, QcCorr, PagCorr, XMLCorr, GraphCorr, TplCorr, [CeOffShore], PrePrintJrnl, ML3G, [CeReviewReq],  [JrnlInst], PeApproval, PeApprovalWait, PE_ApprovalRecDate, JBM_sProof, [RevDely], [ManusType], [InputFileName], [jStyle], TAT , submissiontype, JBM_ProofType, ceduedate, ceoes,ArtIds,CorrFigs, cadAid, CCEHighSpeedWF,JBM_CeRequired, ESA_Status, SiteID, DocketNo, LTH_WF, ChapterID, DOI, Iss, JBM_ID, Auto_WF_Skip, ESA_ArtStatus,JBM_ProcessType  from (";

                //if (Session["sCustAcc"].ToString() != "TF")
                //    strQuery += " Select e.AllocatedTo, e.AllocatedtoName as AssignedTo,'' as ManuscriptID,  e.AutoArtID, e.AllocatedDate, e.Pages, e.Status, e.DeptCode, e.iss as VolIss,  e.RemStatus as InternalRemarks, " + "(Case e.Stage when 'FP' then 'PP' else e.Stage end) as Stage, e.AllocatedBy, e.AllocatedByName, e.SplitNo," + "(Select g.PlatformDesc from " + Init_Tables.gTblPlatform + " g, " + Init_Tables.gTblJrnlInfo + " f where e.AutoArtID=f.JBM_AutoID and f.JBM_Platform=g.PlatformID and f.JBM_Disabled=0) as PlatformDesc," + "d.JBM_ID as ChapterID,NULL as DOI,e.Iss,d.JBM_ID as JBM_ID,d.JBM_Intrnl as JBM_Intrnl," + "(Select g.CustSN from " + Init_Tables.gTblCustomerMaster + " g, " + Init_Tables.gTblJrnlInfo + " f where e.AutoArtID=f.JBM_AutoID and f.CustID=g.CustID and g.Cust_Disabled is NULL and f.JBM_Disabled=0)  as CustSN," + "dbo.fn_ArtStageType(b.ArtStageTypeID) as [CurrentStage], " + "SUBSTRING(b.ArtStageTypeID,1,35) as ArtStageTypeID, " + "Null as GraphArtStageTypeID,Null as CeArtStageTypeID,Null as XMLArtStageTypeID,Null as CeParallel,Null as XMLParallel,Null as NumofFigures,Null as IntrnlID,b.Rev_Wf as WfCode,b.CurrentProcess as CurrentProcess,Null as InitialCE, b.JBM_AutoID, d.MyPet, b.DueDate, b.TeCorr, b.QcCorr, b.PagCorr, b.XMLCorr, b.GraphCorr, b.TplCorr, d.JBM_CeRequiredOffShore as [CeOffShore], d.PrePrintJrnl, d.ML3G, d.JBM_CeReviewRequired as [CeReviewReq], " + "(Select count(*) from "  + Session["sCustAcc"].ToString() + Init_Tables.gTblSplInstructions + " s where s.AutoArtID=d.JBM_AutoID and (s.Stage='All' or (charindex(s.Stage,e.Stage) <> 0)) and (DeptCode Like '%|" + Session["DeptCode"].ToString() + "|%')) as [JrnlInst], Null as PeApproval, Null as PeApprovalWait, Null as PE_ApprovalRecDate, d.JBM_sProof, d.DelyID as [RevDely], NULL as [ManusType], b.InputZipFileName as [InputFileName], d.JTF_S2008 as [jStyle], (Select TM.TAT from JBM_TAT_Master TM where d.JBM_AutoID=TM.JBM_AutoID and TM.Stage='FP' and TM.Priority='N') as TAT " + ", b.submissiontype,d.JBM_ProofType,'' as ceduedate, '' as ceoes,e.ArtIds,b.CorrFigs,NULL as cadAid,NULL as CCEHighSpeedWF,d.JBM_CeRequired, '' as ESA_Status,'' as SiteID, NULL as DocketNo,d.LTH_WF, (case when d.auto_wf_skip is null then '' else 'yes' end) as [Auto_WF_Skip], '' as ESA_ArtStatus,d.JBM_ProcessType as JBM_ProcessType  from " + Session["sCustAcc"].ToString()  + Init_Tables.gTblJBM_Allocation + " e, " + Session["sCustAcc"].ToString()  + Init_Tables.gTblIssueInfo + " b, " + Init_Tables.gTblJrnlInfo + " d where e.AllocatedTo='" + Session["EmpName"].ToString() + "' and e.Status='1' and e.DeptCode='" + Session["DeptCode"].ToString() + "' and e.AutoArtID Like '" + Session["sCustAcc"].ToString() + "%' " + strDispFilterQry + " and b.JBM_autoID=e.AutoArtID  and d.JBM_Disabled='0' and e.Iss=b.iss and e.Stage=b.RevFinStage and d.JBM_AutoID Like '" + Session["sCustAcc"].ToString() + "%' and d.JBM_AutoID=b.JBM_AutoID " + strFilterResult + strArtLocFilter + " UNION ";

                //strQuery += " Select e.AllocatedTo, e.AllocatedtoName as AssignedTo,a.ManuscriptID,  e.AutoArtID, e.AllocatedDate, e.Pages, e.Status, e.DeptCode, e.iss as VolIss,  e.RemStatus as InternalRemarks, " + "(Case e.Stage when 'FP' then 'PP' else e.Stage end) as Stage, e.AllocatedBy, e.AllocatedByName as AllocatedByName, e.SplitNo," + "(Select g.PlatformDesc from " + Init_Tables.gTblPlatform + " g, " + Init_Tables.gTblJrnlInfo + " f where e.AutoArtID=f.JBM_AutoID and f.JBM_Platform=g.PlatformID and f.JBM_Disabled=0) as PlatformDesc," + "a.ChapterID as ChapterID,a.DOI as DOI,a.iss as Iss," + "d.JBM_ID as JBM_ID," + "d.JBM_Intrnl as JBM_Intrnl," + "(Select g.CustSN from " + Init_Tables.gTblCustomerMaster + " g, " + Init_Tables.gTblJrnlInfo + " f, " + Session["sCustAcc"].ToString()  + Init_Tables.gTblChapterOrArticleInfo + " a where a.AutoArtID=e.AutoArtID and a.JBM_AutoID=f.JBM_AutoID and f.CustID=g.CustID and g.Cust_Disabled is NULL and f.JBM_Disabled=0) as CustSN," + "dbo.fn_ArtStageType(b.ArtStageTypeID) as [CurrentStage], SUBSTRING(b.ArtStageTypeID,1,35) as ArtStageTypeID, " + "b.GraphArtStageTypeID as GraphArtStageTypeID, b.CeArtStageTypeID as CeArtStageTypeID," + "b.XMLArtStageTypeID as XMLArtStageTypeID,b.CeParallel as CeParallel," + "b.XMLParallel as XMLParallel,a.NumofFigures as NumofFigures,a.IntrnlID as IntrnlID," + "b.Rev_Wf as WfCode,b.CurrentProcess as CurrentProcess, " + "b.Copyeditor as InitialCE, a.JBM_AutoID, d.MyPet, b.DueDate, b.TeCorr, b.QcCorr, b.PagCorr, b.XMLCorr, b.GraphCorr, b.TplCorr, d.JBM_CeRequiredOffShore as [CeOffShore], d.PrePrintJrnl, d.ML3G, d.JBM_CeReviewRequired as [CeReviewReq], " + "(Select count(*) from " + Session["sCustAcc"].ToString() + Init_Tables.gTblSplInstructions + " s where s.AutoArtID=a.JBM_AutoID and (s.Stage='All' or (charindex(s.Stage,e.Stage) <> 0)) and (DeptCode Like '%|" + Session["DeptCode"].ToString() + "|%')) as [JrnlInst], d.pe_approval as PeApproval, d.pe_approvalwait as PeApprovalWait, b.PE_ApprovalRecDate, d.JBM_sProof, d.DelyID as [RevDely], a.ManusType as [ManusType], b.InputZipFileName as [InputFileName], d.JTF_S2008 as [jStyle], (Select TM.TAT from JBM_TAT_Master TM where d.JBM_AutoID=TM.JBM_AutoID and TM.Stage='FP' and TM.Priority='N') as TAT " + ", b.submissiontype,d.JBM_ProofType,ceduedate, a.ceoes,e.ArtIds,b.CorrFigs,a.cadAid,b.CCEHighSpeedWF,d.JBM_CeRequired, d.ESA_Status, d.SiteID as SiteID,a.DocketNo,d.LTH_WF, (case when d.auto_wf_skip is null then '' else 'yes' end) as [Auto_WF_Skip],b.ESA_Status as ESA_ArtStatus ,d.JBM_ProcessType as JBM_ProcessType   from " + Session["sCustAcc"].ToString()  + Init_Tables.gTblJBM_Allocation + " e, " + Session["sCustAcc"].ToString()  + Init_Tables.gTblStageInfo + " b, " + Session["sCustAcc"].ToString() +  Init_Tables.gTblChapterOrArticleInfo + " a, " + Init_Tables.gTblJrnlInfo + " d where e.AllocatedTo='" + Session["EmpName"].ToString() + "' and e.Status='1' and e.DeptCode='" + Session["DeptCode"].ToString() + "'   " + strDispFilterQry + " and a.AutoArtID=b.AutoArtID and b.AutoArtID=e.AutoArtID and a.JBM_AutoID=d.JBM_AutoID and b.RevFinStage=e.Stage and (substring(e.AutoArtID, 2,1) Like '[0-9]') and a.WIP=1 and d.JBM_Disabled='0' and a.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' and d.JBM_AutoID like '" + Session["sCustAcc"].ToString() + "%' " + strFilterResult + strArtLocFilter + ") TempTable Order By AllocatedDate, (Select Case Substring(stage, 1,3) when 'Fin' then 'zz' when 'Dum' then 'yy' when 'uDum' then 'xx' else Substring(stage, 1,3) end) Desc";




                DataSet ds = new DataSet();

                //strQuery = "Select a.AutoartID as [Job ID],a.IntrnlID as [Internal ID],j.JBM_Intrnl as [Journal ID],a.ChapterID as [Article ID],a.Iss,a.NumofMSP as [Pages],s.revfinstage as [Stage],Convert(varchar, s.DueDate, 106) As DueDate,Convert(varchar, al.allocateddate, 106) As AllocatedDate,CM.CustName,CM.CustSN, p.PlatformDesc,a.DOI, Convert(varchar, s.receiveddate, 106) As ReceivedDate , s.artstagetypeid,a.jbm_autoid,s.rev_wf as WF,j.jbm_mPe,j.jbm_id,j.JBM_WFHCelevel, jbm_ceqc,Smt_Ed,a.NumofFigures,a.NumofTables,CadAid, s.CCEHighSpeedWF,j.SiteID,s.ArtPickup_Loc, s.Ceparallel, s.XMLParallel, s.CurrentProcess,s.GraphArtStageTypeID, s.CeArtStageTypeID,s.CurrentProcess as [RTPStatus], s.CurrentStatus, s.InputZipFileName,CM.CustGroup  from TF_stageinfo s join TF_ArticleInfo a  On a.autoartid=s.autoartid join JBM_Info j On a.jbm_autoid=j.jbm_autoid join JBM_CustomerMaster CM on j.CustID=CM.CustID join TF_Allocation al on al.AutoartID=s.AutoartID join JBM_Platform p on p.platformid=a.platformid where s.dispatchdate Is null and  (s.ArtStageTypeID like '%_S1009_%' ) And s.RevFinStage = 'PRE1' and a.AutoArtID in(Select AutoArtID from TF_Allocation e where a.AutoArtID=e.AutoArtID and isnull(a.Iss, '9999999')= isnull(e.Iss,'9999999') and e.DeptCode in (30,260) and s.revfinstage=e.Stage and e.AllocatedTo='E000002' and e.Status=1)";
                string strQuery = "Select a.AutoartID as [Job ID],a.IntrnlID as [Internal ID],j.JBM_Intrnl as [Journal ID],a.ChapterID as [Article ID],a.Iss,a.NumofMSP as [Pages],s.revfinstage as [Stage],Convert(varchar, s.DueDate, 106) As DueDate,Convert(varchar, al.allocateddate, 106) As AllocatedDate,CM.CustName,CM.CustSN, p.PlatformDesc,a.DOI, Convert(varchar, s.receiveddate, 106) As ReceivedDate , s.artstagetypeid,a.jbm_autoid,s.rev_wf as WF,j.jbm_mPe,j.jbm_id,j.JBM_WFHCelevel, jbm_ceqc,Smt_Ed,a.NumofFigures,a.NumofTables,CadAid, s.CCEHighSpeedWF,j.SiteID,s.ArtPickup_Loc, s.Ceparallel, s.XMLParallel, s.CurrentProcess,s.GraphArtStageTypeID, s.CeArtStageTypeID,s.CurrentProcess as [RTPStatus], s.CurrentStatus, s.InputZipFileName,CM.CustGroup  from " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " s join " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a  On a.autoartid=s.autoartid join " + Init_Tables.gTblJrnlInfo + " j On a.jbm_autoid=j.jbm_autoid join " + Init_Tables.gTblCustomerMaster + " CM on j.CustID=CM.CustID join " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " al on al.AutoartID=s.AutoartID join " + Init_Tables.gTblPlatform + " p on p.platformid=a.platformid where s.dispatchdate Is null " + strFilterResult + " And s.RevFinStage = al.stage and a.AutoArtID in(Select AutoArtID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " e where a.AutoArtID=e.AutoArtID and isnull(a.Iss, '9999999')= isnull(e.Iss,'9999999') and e.DeptCode in (" + Session["DeptCode"].ToString() + ") and s.revfinstage=e.Stage and al.AllocatedTo='" + Session["EmpAutoId"].ToString() + "' and e.Status=1 " + strArtLocFilter + strDispFilterQry + ")";  //and  (s.ArtStageTypeID like '%_S1009_%' ) 

                if (Session["sConnSiteDB"].ToString() != null)
                {
                    ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {a[0].ToString(),a[2].ToString(),a[3].ToString(), a[1].ToString(), a[7].ToString(),a[8].ToString(),a[4].ToString(), a[5].ToString(),a[6].ToString(),a[11].ToString(), CreateWebControl(a, "Pickup"),
                                            CreateWebControl(a, "check")
                                         };

                    return Json(new { dataDB = JSONString }, JsonRequestBehavior.AllowGet);
                }
                else {
                    return Json(new { dataDB = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                
            }
            catch (Exception)
            {

                return Json(new { dataDB = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult ChangeDepartment(string id)
        {
            try
            {
                Session["DeptCode"] = id;
                return Json(new { dataDB = "Success" }, JsonRequestBehavior.AllowGet);
                //return RedirectToAction("GetInboxAlloatedDetails");
               
            }
            catch (Exception)
            {
                return Json(new { dataEmp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        public async Task<ActionResult> RemoveAllottedJobs(string CheckedArticleID)
        {

            try
            {
                List<string> removeIdsColl = JsonConvert.DeserializeObject<List<string>>(CheckedArticleID);
                if (removeIdsColl.Count > 0)
                {
                    for (int i = 0; i < removeIdsColl.Count; i++)
                    {
                        
                        string sJobID = ""; string sInternalID = ""; string sArticleID = ""; string sStage = ""; string sPages = ""; string sArtStageTypeID = ""; string sCurrentProcess = ""; string sJBMAutoID = ""; string sIss = "";
                        sJobID = removeIdsColl[i].Split('|')[0].ToString();
                        sInternalID = removeIdsColl[i].Split('|')[1].ToString();
                        sArticleID = removeIdsColl[i].Split('|')[2].ToString();
                        sStage = removeIdsColl[i].Split('|')[3].ToString();
                        sPages = removeIdsColl[i].Split('|')[4].ToString();
                        sArtStageTypeID = removeIdsColl[i].Split('|')[5].ToString();
                        sCurrentProcess = removeIdsColl[i].Split('|')[6].ToString();
                        sJBMAutoID = removeIdsColl[i].Split('|')[7].ToString();
                        sIss = removeIdsColl[i].Split('|')[8].ToString();

                        string strVolIssNo = "";
                        if (strVolIssNo != "")
                            strVolIssNo = " and Iss='" + sIss + "'";
                        else
                            strVolIssNo = " and Iss Is Null ";

                        //To update allocation table
                        string strResult = DBProc.GetResultasString("Delete from " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " where AllocatedTo='" + Session["EmpAutoId"].ToString() + "' and AutoArtID='" + sJobID + "' and Stage='" + sStage + "' and DeptCode='" + Session["DeptCode"].ToString() + "'" + strVolIssNo, Session["sConnSiteDB"].ToString());

                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataDB = "Success"}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataDB = "Failed"}, JsonRequestBehavior.AllowGet);
            }   
            
            //return View();
        }

        public ActionResult GetIncomingJobs()
        {
            try
            {
                string strSearchQry = "";
                DataSet ds = new DataSet();

                string strQuery = ""; string strcadaid = "a.ChapterID as [ArticleID]"; string strArtLocFilter = "";

                if (Session["sCustAcc"].ToString() == "TF" & (Request.Url.AbsoluteUri.ToString().ToLower().Contains("10.18") | Request.Url.AbsoluteUri.ToString().ToLower().Contains("cps.kwglobal.com") | Request.Url.AbsoluteUri.ToString().ToLower().Contains("smarttrack.cenveo.com/smarttrack-ch")))
                    strArtLocFilter = " and (d.SiteID='" + Init_SiteID.gSiteChennai + "' OR RIGHT(b.ArtPickup_Loc, 5)= '" + Init_SiteID.gSiteChennai + "' ) ";
                else if (Session["sCustAcc"].ToString() == "TF" & (Request.Url.AbsoluteUri.ToString().Contains("10.21.") | Request.Url.AbsoluteUri.ToString().ToLower().Contains("cpsblr.cenveo.com") | Request.Url.AbsoluteUri.ToString().ToLower().Contains("localhost")))
                {
                    strArtLocFilter = " and (d.SiteID='" + Init_SiteID.gSiteBangalore + "' OR RIGHT(b.ArtPickup_Loc, 5)= '" + Init_SiteID.gSiteBangalore + "' ) ";
                }

                if (Session["DeptCode"].ToString() == Init_Department.gDcGraphics | Session["DeptCode"].ToString() == Init_Department.gDcInward | Session["DeptCode"].ToString() == Init_Department.gDcImInward)
                {
                    if (Session["DeptCode"].ToString() == Init_Department.gDcGraphics)
                    {
                        strQuery = "Select  *  from ( ";

                        if (Session["sCustAcc"].ToString() != "TF")
                        {
                            strQuery += "Select b.ArtStageTypeID, d.JBM_AutoID as AutoArtID, Null [IntrnlID], NULL as [ArticleID], NULL as DOI, d.JBM_Intrnl as [JID], b.iss as [Iss], b.RevFinStage as [Stage], b.CorrFigs as [Pages],b.CurrentProcess as CP, convert(varchar(12), b.ReceivedDate, 106) as [Received Date], convert(varchar(12), b.loginentereddate, 106) as [Login Entereddate], convert(varchar(12), b.DueDate, 106) as [Due Date],'' as CeReceivedDate,'' as [CE Due Date], '' as [XMLCorr], '' as [TECorr], '' as [PagCorr], '' as [QCCorr],b.GraphCorr as [GraphCorr], d.JBM_AutoID, b.GrapStatus,NULL as NumofFigures,b.DueDate [DDate],'' as [CEDate],NULL as [CCEHighSpeedWF],d.JBM_CeRequiredOffShore as [CeOffShore]  , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], b.ArtPickup_Loc,d.LTH_WF,'' as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType,'' as [Last Acc. By]     from  " + Session["sCustAcc"].ToString() + "_IssueInfo b, " + Init_Tables.gTblJrnlInfo + " d where b.JBM_AutoID=d.JBM_AutoID and  d.JBM_Disabled='0' and " + Session["sCustAcc"].ToString() == "JW" ? " Composition=1 and " : " " + "  b.ReceivedDate is not null " + " and b.DispatchDate is Null and b.GrapStatus is null and ((b.CorrFigs is not null and b.CorrFigs <> '0'))  and b.iss is not null " + strArtLocFilter + " union ";
                        }

                        strQuery += "Select b.ArtStageTypeID, a.AutoArtID, a.IntrnlID as [IntrnlID]," + strcadaid + ", a.DOI, d.JBM_Intrnl as [JID], a.iss as [Iss], b.RevFinStage as [Stage], (case when b.RevFinStage='FP' then  a.NumofFigures else b.CorrFigs end) as [Pages],b.CurrentProcess as CP, " + "convert(varchar(12), b.ReceivedDate, 106) as [Received Date],convert(varchar(12), b.loginentereddate, 106) as [Login Entereddate], convert(varchar(12), b.DueDate, 106) as [Due Date],'' as CeReceivedDate,'' as [CE Due Date], '' as [XMLCorr], '' as [TECorr], '' as [PagCorr], '' as [QCCorr],b.GraphCorr as [GraphCorr], d.JBM_AutoID, b.GrapStatus, a.NumofFigures,b.DueDate [DDate],'' as [CEDate],b.CCEHighSpeedWF as [CCEHighSpeedWF],d.JBM_CeRequiredOffShore as [CeOffShore] , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], b.ArtPickup_Loc,d.LTH_WF,(Select EmbargoDate From " + Session["sCustAcc"].ToString() + "_ProdInfo K where k.Autoartid=a.AutoArtID and (k.EmbargoDate is not NULL And k.EmbargoDate<>'')) as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType,'' as [Last Acc. By]  " + "from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a, " + Session["sCustAcc"].ToString() + "_StageInfo b, " + Init_Tables.gTblJrnlInfo + " d where a.AutoArtID=b.AutoArtID and a.JBM_AutoID=d.JBM_AutoID and a.WIP='1' and d.JBM_Disabled='0' and " + Session["sCustAcc"].ToString() == "JW" ? " Composition=1 and " : " " + " b.ReceivedDate is not null and b.DispatchDate is Null and b.GrapStatus is null and ((a.NumofFigures is not null and a.NumofFigures <> '0') or (b.CorrFigs is not null and b.CorrFigs <> '0'))  " + strArtLocFilter + "" + ") TempTable where TempTable.Pages is not null and TempTable.AutoArtID not in(Select AutoArtID from " + Session["sCustAcc"].ToString() + "_Allocation e where TempTable.AutoArtID=e.AutoArtID and isnull(TempTable.Iss, '9999999')= isnull(e.Iss,'9999999') and TempTable.Stage=e.Stage and e.DeptCode='" + Init_Department.gDcGraphics + "') order by [Due Date]";
                    }
                }
                else if (Session["DeptCode"].ToString() == Init_Department.gDcInward | Session["DeptCode"].ToString() == Init_Department.gDcImInward)
                {
                    string strJobFilter = "(b.ArtStageTypeID like '%_S1001_%' and b.ReceivedDate is not null and b.DispatchDate is null)";
                    //if (SessionHandler.gJwAccItm.Contains("|IDW|") & Session["sSiteID"].ToString() == Init_SiteID.gSiteChennai)
                    if (Session["sSiteID"].ToString() == Init_SiteID.gSiteChennai)
                        strSearchQry += " and JBM_IM<>1 ";

                    if (Session["DeptCode"].ToString() == Init_Department.gDcImInward)
                        strJobFilter = "(b.ArtStageTypeID like '%_S1070_%' and b.ArtStageTypeID not like '%_S1071_S1070_%' and ((b.ReceivedDate is null and b.DispatchDate is null) or (b.ReceivedDate is not null and b.DispatchDate is null)))";

                    strQuery = "Select *  from ( ";

                    if (Session["sCustAcc"].ToString() != "TF")
                        strQuery += "Select b.ArtStageTypeID, d.JBM_AutoID as AutoArtID, Null [IntrnlID], NULL as [ArticleID], NULL as DOI, d.JBM_Intrnl as [JID], b.iss as [Iss], b.RevFinStage as Stage, NULL as [Pages],b.CurrentProcess as CP, convert(varchar(12), b.ReceivedDate, 106) as [Received Date], convert(varchar(12), b.loginentereddate, 106) as [Login Entereddate],'' as CeReceivedDate, convert(varchar(12), b.DueDate, 106) as [Due Date], '' as [CE Due Date], '' as [XMLCorr], '' as [TECorr], '' as [PagCorr], '' as [QCCorr],b.GraphCorr as [GraphCorr], d.JBM_AutoID, b.ReceivedDate as [RecDate], Null as GrapStatus, Null as NumofFigures,'' as [CEDate],b.DueDate as [DDate],NULL as [CCEHighSpeedWF],d.JBM_CeRequiredOffShore as [CeOffShore] , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], b.ArtPickup_Loc,'' as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType,'' as [Last Acc. By]   from " + Session["sCustAcc"].ToString() + "_IssueInfo b, " + Init_Tables.gTblJrnlInfo + " d where b.JBM_AutoID=d.JBM_AutoID and d.JBM_Disabled='0' and b.ArtStageTypeID like '%_S1001_%' and b.ReceivedDate is not null and b.DispatchDate is null " + strArtLocFilter + " UNION ";

                    strQuery += "Select b.ArtStageTypeID, a.AutoArtID, a.IntrnlID as [IntrnlID]," + strcadaid + ", a.DOI, d.JBM_Intrnl as [JID], a.iss as [Iss], b.RevFinStage as Stage, a.NumofFigures as [Pages],b.CurrentProcess as CP, " + "convert(varchar(12), b.ReceivedDate, 106) as [Received Date],convert(varchar(12), b.loginentereddate, 106) as [Login Entereddate],convert(varchar(12), b.DueDate, 106) as [Due Date],convert(varchar(12), b.CeRecDate, 106) as CeReceivedDate, convert(varchar(12), b.CeDueDate, 106) as [CE Due Date], '' as [XMLCorr], '' as [TECorr], '' as [PagCorr], '' as [QCCorr],b.GraphCorr as [GraphCorr], d.JBM_AutoID, b.ReceivedDate as [RecDate], b.GrapStatus, a.NumofFigures,b.CeDueDate as [CEDate],b.DueDate as [DDate],b.CCEHighSpeedWF as [CCEHighSpeedWF],NUll as [CeOffShore] , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], b.ArtPickup_Loc,(Select EmbargoDate From " + Session["sCustAcc"].ToString() + "_ProdInfo K where k.Autoartid=a.AutoArtID and (k.EmbargoDate is not NULL And k.EmbargoDate<>'')) as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType,'' as [Last Acc. By]    " + "from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo +" a, " + Session["sCustAcc"].ToString() + "_StageInfo b, " + Init_Tables.gTblJrnlInfo + " d where a.AutoArtID=b.AutoArtID and a.JBM_AutoID=d.JBM_AutoID and a.WIP='1' and d.JBM_Disabled='0' " + strSearchQry + " and (" + strJobFilter + " )) TempTable where TempTable.AutoArtID Not In(Select AutoArtID from " + Session["sCustAcc"].ToString() + "_Allocation e where TempTable.AutoArtID=e.AutoArtID And isnull(TempTable.Iss, '9999999')= isnull(e.Iss,'9999999') and TempTable.Stage=e.Stage and DeptCode='" + Session["DeptCode"].ToString() + "')  order by [RecDate]";

                }
                else {

                    string strFilterResult = "";
                    string strMeFilterResult = "";
                    string strFailureFilter = "";
                    string strIssFailureFilter = "";
                    string strMeIssFilterResult = "";
                    string strCEReviewFilterResult = "";
                    string StrOnshoreMeFilterRes = "";
                    string strEmpArtStage = "";
                    string ReMail = "";
                    string StrCeQcFilter = "";

                    switch (Session["DeptCode"].ToString())
                    {
                        case Init_Department.gDcInward:
                            {
                                strFilterResult = " like '%_S1001_%'";
                                strEmpArtStage = Init_Barcode.gstrLogin;
                                break;
                            }

                        case Init_Department.gDcRE:
                            {
                                strFilterResult = " like '%_S1002_%'";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1002_%') "; 
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1002_%') ";

                                if (Session["sSiteID"].ToString() == Init_SiteID.gSiteMumbai)
                                    ReMail = " and d.ReMail=1 ";
                                else
                                    ReMail = " and (d.ReMail is Null or d.ReMail=0) ";
                                strEmpArtStage = Init_Barcode.gstrCeRapid;
                                break;
                            }

                        case  Init_Department.gDcCE:
                            {
                                strFilterResult = " like '%_S1003_%'";
                                strMeFilterResult = " OR (b.ArtStageTypeID like '%_S1152_%' AND b.CurrentProcess='S1152') Or b.ArtStageTypeID like '%_S1161_S1002_%' ";
                                StrOnshoreMeFilterRes = " OR (i.ArtStageTypeID like '%_S1187_%' AND i.CurrentProcess='S1187') ";
                                strCEReviewFilterResult = " OR (b.ArtStageTypeID like '%_S1038_%' ) ";
                                strMeIssFilterResult = " OR (i.ArtStageTypeID like '%_S1152_%' AND i.CurrentProcess='S1152') Or i.ArtStageTypeID like '%_S1161_S1002_%' ";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1003_%') ";
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1003_%') ";
                                StrCeQcFilter = " OR (b.ArtStageTypeID like '%_S1057_S1057_%') ";
                                strEmpArtStage = Init_Barcode.gstrCopyEditing;
                                break;
                            }

                        case Init_Department.gDcXML:
                            {
                                strFilterResult = " like '%_S1004_%'";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1004_%') ";
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1004_%') ";
                                strEmpArtStage = Init_Barcode.gstrXML;
                                break;
                            }

                        case Init_Department.gDcPag:
                            {
                                strFilterResult = " like '%_S1005_%'";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1005_%') ";
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1005_%') ";
                                strEmpArtStage = Init_Barcode.gstrPagination;
                                break;
                            }

                        case  Init_Department.gDcTE:
                            {
                                strFilterResult = " like '%_S1006_%'";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1006_%') ";
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1006_%') ";
                                strEmpArtStage = Init_Barcode.gstrTech;
                                break;
                            }

                        case Init_Department.gDcQC:
                            {
                                strFilterResult = " like '%_S1007_%'";
                                strFailureFilter = " OR (b.ArtStageTypeID like '%_S1161_S1007_%') ";
                                strIssFailureFilter = " OR (i.ArtStageTypeID like '%_S1161_S1007_%') ";
                                strEmpArtStage = Init_Barcode.gstrQc;
                                break;
                            }

                        case Init_Department.gDcMyPet:
                            {
                                strFilterResult = " like '%_S1041_%'";
                                strEmpArtStage = Init_Barcode.gstrMyPet;
                                break;
                            }

                        case Init_Department.gDcICE:
                            {
                                strFilterResult = " like '%_S1079_%'";
                                strEmpArtStage = Init_Barcode.gstrICE;
                                break;
                            }

                        case Init_Department.gDcTechSupport:
                            {
                                strFilterResult = " like '%_S1110_%'";
                                strEmpArtStage = Init_Barcode.gstrTechSupport;
                                break;
                            }

                        case Init_Department.gDcConversion:
                            {
                                strFilterResult = " like '%_S1113_%'";
                                strEmpArtStage = Init_Barcode.gstrConversion;
                                break;
                            }

                        case Init_Department.gDcPPC:
                            {
                                strFilterResult = " like '%_S1119_%'";
                                strEmpArtStage = Init_Barcode.gstrPPC;
                                break;
                            }

                        case Init_Department.gDcIm:
                            {
                                strFilterResult = " like '%_S1027_%'";
                                strEmpArtStage = Init_Barcode.gstrIM;
                                break;
                            }

                        case Init_Department.gDcTemplate:
                            {
                                strFilterResult = " like '%_S1024_%'";   
                                strEmpArtStage = Init_Barcode.gstrTemplate;
                                break;
                            }

                        case Init_Department.gDcProjectManagement:
                        case Init_Department.gDcAM:
                            {
                                strFilterResult = " like '%_S1045_%'";
                                strEmpArtStage = Init_Barcode.gstrProjectManagement;
                                break;
                            }

                        case Init_Department.gDcAltText:
                            {
                                strFilterResult = " like '%_S1128_%'";
                                strEmpArtStage = Init_Barcode.gstrArtPass2AltText;
                                break;
                            }
                    }

                    if (Session["sCustAcc"].ToString() != "TF")
                        strFailureFilter = "";

                    strQuery = "Select AutoArtID,[ArticleID],DOI,IntrnlID,JID,JBM_AutoID,Iss,[Received Date],[Login Entereddate],[Due Date],CeReceivedDate,[CE Due Date],Stage,WF,ArtStageTypeID,CP,Pages,[PlatformDesc],ActualPages,[XMLCorr],[PagCorr],[TECorr],[QCCorr],[GraphCorr],GrapStatus,NumofFigures,[CEDate],[DDate],[CCEHighSpeedWF],[CeOffShore],HighPriority,[JrnlSiteID],ArtPickup_Loc,LTH_WF, [EmbargoDate],JBM_ProcessType, [Last Acc. By]  from ( ";

                    if (Session["sCustAcc"].ToString() != "TF")
                        strQuery += "Select i.JBM_AutoId as AutoArtID, d.JBM_Intrnl as [ArticleID],  '' as DOI, '' as IntrnlID, d.JBM_Intrnl as JID, d.JBM_AutoID, i.iss as Iss, " + "convert(varchar(12), i.ReceivedDate,106) as [Received Date],convert(varchar(12), i.loginentereddate,106) as [Login Entereddate], convert(varchar(12), i.DueDate,106) as [Due Date], '' as CeReceivedDate,'' as [CE Due Date], " + "i.RevFinStage as Stage, i.Rev_WF as WF, i.ArtStageTypeID,i.CurrentProcess as CP, " + "(case when (cast(i.RevFinPages as int)) is null then '0' else (cast(i.RevFinPages as int)) end) as Pages, " + "(select p.PlatformDesc from " + Init_Tables.gTblPlatform + " p where p.platformID=d.JBM_Platform) as [PlatformDesc]," + "(Select (case when sum(cast(a.ActualPages as int)) is null then '0' else sum(cast(a.ActualPages as int)) end) from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo +"  a where a.JBM_AutoID=i.JBM_AutoID) as ActualPages," + "(case when i.XMLCorr is null then '0' else i.XMLCorr end) as [XMLCorr], (case when i.PagCorr is null then '0' else i.PagCorr end) as [PagCorr], " + "(case when i.TeCorr is null then '0' else i.TeCorr end) as [TECorr],(case when i.QcCorr is null then '0' else i.QcCorr end) as [QCCorr],(case when i.GraphCorr is null then '0' else i.GraphCorr end) as [GraphCorr], Null as GrapStatus, Null as NumofFigures,'' as [CEDate],i.DueDate as [DDate],NULL as [CCEHighSpeedWF],d.JBM_CeRequiredOffShore as [CeOffShore] , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], i.ArtPickup_Loc,d.LTH_WF,'' as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType, '' as [Last Acc. By]    " + "from  " + Session["sCustAcc"].ToString() + "_IssueInfo i, " + Init_Tables.gTblJrnlInfo + " d where i.DispatchDate is null and i.HoldArtStatus is null and i.DueDate is not null and d.JBM_AutoID=i.JBM_AutoId and d.JBM_Disabled='0' " + ReMail + " and ( i.ArtStageTypeID " + strFilterResult + strMeIssFilterResult + strIssFailureFilter + StrOnshoreMeFilterRes.Replace("b.", "i.")  + StrCeQcFilter.Replace("b.", "i.") + ")" + "" + strArtLocFilter.Replace("b.", "i.") + " Union ";


                    strQuery += "Select a.Autoartid as AutoArtid, " + strcadaid + ",  a.DOI, a.IntrnlID, d.JBM_Intrnl as JID, d.JBM_AutoID, a.iss as Iss, " + "(case when b.RevFinStage='FP' then convert(varchar(12), b.LoginRecDate,106)  else convert(varchar(12), b.ReceivedDate,106) end ) as [Received Date], convert(varchar(12), b.loginentereddate,106) as [Login Entereddate] , " + "convert(varchar(12), b.DueDate, 106) as [Due Date], convert(varchar(12), b.CeRecDate, 106) as CeReceivedDate,convert(varchar(12), b.CeDueDate, 106) as [CE Due Date]," + "b.RevFinStage as Stage,b.Rev_WF as WF,b.ArtStageTypeID,b.CurrentProcess as CP," + "(case b.RevFinStage when 'FP' then ((case when (cast(a.NumofMSP as int)) is null then '0' else (cast(a.NumofMSP as int)) end))else(case when (cast(b.RevFinPages as int)) is null then '0' else (cast(b.RevFinPages as int)) end) end) as Pages," + "p.PlatformDesc, a.ActualPages as ActualPages,(case when b.XMLCorr is null then '0' else b.XMLCorr end) as [XMLCorr], (case when b.PagCorr is null then '0' else b.PagCorr end) as [PagCorr], " + "(case when b.TeCorr is null then '0' else b.TeCorr end) as [TECorr],(case when b.QcCorr is null then '0' else b.QcCorr end) as [QCCorr],(case when b.GraphCorr is null then '0' else b.GraphCorr end) as [GraphCorr], b.GrapStatus, a.NumofFigures,b.CeDueDate as [CEDate],b.DueDate as [DDate],b.CCEHighSpeedWF as [CCEHighSpeedWF],d.JBM_CeRequiredOffShore as [CeOffShore]  , d.Priority as HighPriority, d.SiteID as [JrnlSiteID], b.ArtPickup_Loc,d.LTH_WF,(Select EmbargoDate From " + Session["sCustAcc"].ToString() + "_ProdInfo K where k.Autoartid=a.AutoArtID and (k.EmbargoDate is not NULL And k.EmbargoDate<>'')) as EmbargoDate, d.JBM_ProcessType as JBM_ProcessType, (Select EM.EmpName from " + Init_Tables.gTblEmployee + " EM where EM.EmpAutoID = (Select Top 1 EmpAutoID from " + Session["sCustAcc"].ToString() + "_ProdStatus PS where PS.autoartid = a.AutoArtID and PS.empArtStage = '" + strEmpArtStage + "' and SUBSTRING(PS.ArtStageInfo,0,5) = b.RevFinStage order by EmpOutTime desc)) as [Last Acc. By] " + "from  " + Session["sCustAcc"].ToString() + "_StageInfo b, " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo +" a, " + Init_Tables.gTblJrnlInfo + " d, " + Init_Tables.gTblPlatform + " p " + "where(b.HoldArtStatus is null and b.DispatchDate Is null And b.DueDate Is Not null and a.AutoArtID=b.AutoArtID and d.JBM_AutoID= a.JBM_AutoID and a.WIP='1' and d.JBM_Disabled='0' " + ReMail + " and p.platformID = a.PlatformID)" + "and ( b.ArtStageTypeID " + strFilterResult + strMeFilterResult + strFailureFilter + StrCeQcFilter + StrOnshoreMeFilterRes.Replace("i.", "b.") + strCEReviewFilterResult + ")" + " " + strSearchQry + strArtLocFilter + ") TempTable where (AutoartID not in(Select AutoArtID from " + Session["sCustAcc"].ToString() + "_Allocation e where TempTable.AutoArtID=e.AutoArtID and isnull(TempTable.Iss, '9999999')= isnull(e.Iss,'9999999') and e.DeptCode='" + Session["DeptCode"].ToString() + "' and TempTable.Stage=e.Stage)) ";

                    if (Session["sCustAcc"].ToString() == "TF")
                        strQuery = strQuery + " order by HighPriority desc";
                }
                
                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
               
                
                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                                CreateWebControl(a, "checkIncoming"),  //Checkbox
                                                CreateWebControl(a, "JobIDCol"), //a[0].ToString(), AutoArtID
                                                CreateWebControl(a, "IdentifyHightSpeedCol"),// a[3].ToString(), //InternalID
                                                a[4].ToString(),    //JournalID
                                                a[1].ToString(),    //ArticleID
                                                a[6].ToString(),    //Iss
                                                a[12].ToString(),   //Stage
                                                a[16].ToString(),   //Pages
                                                a[8].ToString(),    //Login Date
                                                a[7].ToString(),    //ReceivedDate
                                                a[9].ToString(),    // Due Date
                                                a[36].ToString()    // Last Acces. By
                                         };

                return Json(new { dataDB = JSONString }, JsonRequestBehavior.AllowGet);


            }
            catch (Exception)
            {
                return Json(new { dataDB = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public string CreateWebControl(DataRow row, string strType)
        {
            
            string formControl = string.Empty;
            try
            {
                
                string uniqueID = row[0].ToString();
                string strJBMAutoID = row[5].ToString();
                string ArtStageTypeID = row[14].ToString();
                string CurrentProcess = row[15].ToString();
                string WFCode = row[16].ToString();
                string strStage = row[12].ToString();

                int corrcount = 0;
                string strGraphStatus = "";
                string strNumofFigs = "";
                bool blnGraphicsNotCompleted = false;
                string StrEmbargoDate = "";

                if (strType == "check")
                {
                    formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "'/>";
                }
                else if (strType == "Pickup")
                {
                    var lstProcess = new List<string>();
                    lstProcess = clsWFProc.getNextProcess(uniqueID, ArtStageTypeID, WFCode, Session["sConnSiteDB"].ToString());
                    string strNextCollection = string.Empty;

                    for (int i = 0; i < lstProcess.Count; i++)
                    {
                        if (i == 0){strNextCollection = lstProcess[i].ToString();}
                        else { strNextCollection += "," + lstProcess[i].ToString(); }
                        
                    }

                    formControl = "<a href='javascript:void(0);'  onClick=\"btnPickupJob('" + uniqueID + "')\" class='btn btn-round btn-outline-secondary btn-sm' id='pickup" + uniqueID + "'  data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-nxt='" + strNextCollection + "' data-wf='" + WFCode + "'>Pickup</a>";
                }
                else if (strType == "checkIncoming")
                {
                    formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "'/>";

                    if (Session["DeptCode"].ToString() == Init_Department.gDcQC)
                    {
                        //QCCorr 
                        corrcount = row[22].ToString() != "" ? Convert.ToInt32(row[22].ToString().Trim()) : 0;
                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcPag)
                    {
                        //PagCorr
                        corrcount = row[20].ToString() != "" ? Convert.ToInt32(row[20].ToString().Trim()) : 0;

                        if (strStage == "FP")
                        {
                            if (row[25].ToString().Trim() != "")  //NumofFigures
                            {
                                strNumofFigs = row[25].ToString().Trim();
                                if (strNumofFigs != "0")
                                {
                                    //GrapStatus
                                    strGraphStatus = row[24].ToString().Trim();

                                    if (strGraphStatus != "1")
                                        blnGraphicsNotCompleted = true;
                                }
                            }
                        }

                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcTE)
                    {
                        //TECorr
                        corrcount = row[21].ToString() != "" ? Convert.ToInt32(row[21].ToString().Trim()) : 0;
                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcXML)
                    {
                        //XMLCorr
                        corrcount = row[19].ToString() != "" ? Convert.ToInt32(row[19].ToString().Trim()) : 0;
                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcGraphics)
                    {
                        //GraphCorr
                        corrcount = row[23].ToString() != "" ? Convert.ToInt32(row[23].ToString().Trim()) : 0;
                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcRE)
                    {
                        corrcount = 0;
                        if (row[14].ToString().Contains("S1105_S1010_S1002_") == true)
                        {
                            //Job ID COlumn        
                            //formControl = "<div style='color:pink' tooltip='ESA Completed to RE'>" + row[0].ToString().Trim() + "</div>";
                            formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='pink' data-tt='ESA Completed to RE'/>";

                        }
                        else if (row[14].ToString().Contains("S1104_S1010_S1002_") == true)
                        {
                            //formControl = "<div style='color:HotPink' tooltip='CleanUp Completed to RE'>" + row[0].ToString().Trim() + "</div>";
                            formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='hotpink' data-tt='CleanUp Completed to RE'/>";
                        }
                    }
                    else if (Session["DeptCode"].ToString() == Init_Department.gDcCE)
                    {
                        string strJbmProcesstype = string.Empty;
                        if (row[36].ToString() != "") //JBM_ProcessType
                            strJbmProcesstype = row[36].ToString();

                        if (strJbmProcesstype == "RTP" & strStage == "FP")
                        {
                            if (row[25].ToString() != "")  //NumofFigures
                            {
                                strNumofFigs = row[25].ToString();
                                if (strNumofFigs != "0")
                                {
                                    if (row[24].ToString() != "")  //GrapStatus
                                        strGraphStatus = row[24].ToString();

                                    if (strGraphStatus != "1")
                                        blnGraphicsNotCompleted = true;
                                }
                            }
                        }
                    }

                    if (strStage == "FP")
                    {
                        if (row[35].ToString() != "") //EmbargoDate
                            StrEmbargoDate = row[35].ToString();
                    }


                    if (blnGraphicsNotCompleted)
                    {
                        //formControl = "<div' style='color:Black;background:gray;'>" + row[5].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='black' data-bg='gray' tooltip='Graphics not completed'/>";
                        //CheckBox chkSel = (CheckBox)e.Row.FindControl("chkSel");
                        //chkSel.Enabled = false;
                        //chkSel.ToolTip = "Graphics not completed";
                    }
                    else if (row[15].ToString().Trim() != "" && row[15].ToString().Trim() == "S1057" | row[15].ToString().Trim() == "S1058")
                        //formControl = "<div' style='color:Black;background:LightGray;'>" + row[0].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='black' data-bg='lightgray'/>";
                    else if (corrcount == 0)
                    {
                        //formControl = "<div' style='color:Black;background:AliceBlue;'>" + row[0].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='black' data-bg='aliceBlue'/>";
                    }
                    else if (corrcount == 1)
                    {
                        //formControl = "<div' style='color:Black;background:DeepSkyBlue;'>" + row[0].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='white' data-bg='#20c997'/>";
                    }
                    else if (corrcount == 2)
                    {
                        //formControl = "<div' style='color:Black;background:CornflowerBlue;'>" + row[0].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='white' data-bg='#28a745'/>";
                    }
                    else if (corrcount >= 3)
                    {
                        //formControl = "<div' style='color:Black;background:OrangeRed;'>" + row[0].ToString().Trim() + "</span>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='white' data-bg='#ed1e26'/>";
                    }

                    if (StrEmbargoDate != "" & Session["sCustAcc"].ToString() == "EH")
                    {
                        //formControl = "<div style='color:HotPink;background:yellow;' tooltip='Embargo Article'>" + row[0].ToString().Trim() + "</div>";
                        formControl = "<input type='checkbox'  onClick=\"chkIncomingJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-at='" + strJBMAutoID + "' data-st='" + ArtStageTypeID + "' data-cp='" + CurrentProcess + "' data-clr='HotPink' data-bg='yellow'  tooltip='Embargo Article'/>";
                    }

                }
                else if (strType == "IdentifyHightSpeedCol")
                {
                    if (row[29].ToString() != "") //CCEHighSpeedWF
                    {
                        formControl = "<span' title='HighSpeed Flow'>" + row[3].ToString().Trim() + "</span>&nbsp;<img style='padding-right=3px' src='../Images/10_flag_purple.png' />";
                    }
                    else
                    {
                        formControl = "<span' title='Normal Flow'>" + row[3].ToString().Trim() + "</span>";
                    }
                }
                else if (strType == "JobIDCol")
                {
                    if (row[31].ToString() != "") //HighPriority
                    {
                        formControl = "<span' title='High Priority'>" + row[0].ToString().Trim() + "</span>&nbsp;<img style='padding-right=3px' src='../Images/high_priority_flag.png' />";
                    }
                    else
                    {
                        formControl = "<span' title='Normal'>" + row[0].ToString().Trim() + "</span>";
                    }
                }
                

                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
      
        public ActionResult AllotJobsInbox(string CheckedArticleID)
        {
            try
            {

                List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedArticleID);
                if (chkIds.Count > 0)
                {
                    for (int i = 0; i < chkIds.Count; i++)
                    {
                        string strAutoArtID = chkIds[i].Split('|')[0].ToString().Trim();
                        string strInternalID = chkIds[i].Split('|')[1].ToString().Trim();
                        string strArticleID = chkIds[i].Split('|')[2].ToString().Trim();
                        string strStage = chkIds[i].Split('|')[3].ToString().Trim();
                        string strPages = chkIds[i].Split('|')[4].ToString().Trim();
                        string strArtStageTypeID = chkIds[i].Split('|')[5].ToString().Trim();
                        string strCurrentProcess = chkIds[i].Split('|')[6].ToString().Trim();
                        string strJAutoID = chkIds[i].Split('|')[7].ToString().Trim();
                        string strIss = chkIds[i].Split('|')[8].ToString().Trim();
                        string strAllocatedTo = Session["EmpAutoId"].ToString();
                        string strAllocatedDate= DateTime.Now.ToString("dd-MMM-yy");
                        string strAllocatedBy = Session["EmpAutoId"].ToString();
                        string strAllocatedByName = Session["EmpName"].ToString();
                        string strAllocatedToName = Session["EmpName"].ToString();

                        if (strIss == "")
                        {
                            strIss = "NULL";
                        }

                        string strpmcurrentprocess = "";
                        string strpmgrapharttypeid = "";

                        if (strCurrentProcess != "")
                        {
                            strpmcurrentprocess = strCurrentProcess;
                        }
                        if (strArtStageTypeID.Contains("_S1144_") && strArtStageTypeID.Contains("S1144"))
                        {
                            bool result = DBProc.UpdateRecord("Update " + Session["sCustAcc"].ToString() + "_StageInfo Set currentprocess='S1038' where AutoArtID='" + strAutoArtID + "' and RevfinStage='" + strStage + "'", Session["sConnSiteDB"].ToString());
                        }

                        if (Session["DeptCode"].ToString() == Init_Department.gDcProjectManagement)
                        {
                            strArtStageTypeID = strpmgrapharttypeid;
                        }

                        string strRemarksFlag = "NULL";

                        DataSet ds = new DataSet();
                        ds = DBProc.GetResultasDataSet("Select AllocatedToName from " + Session["sCustAcc"].ToString() + "_Allocation where AutoArtID='" + strAutoArtID + "' and DeptCode='" + Session["DeptCode"].ToString() + "' and Stage='" + strStage + "'", Session["sConnSiteDB"].ToString());
                        if (ds.Tables[0].Rows.Count == 0)
                        {
                            ds = new DataSet();
                            ds = DBProc.GetResultasDataSet("Select AutoArtId from " + Session["sCustAcc"].ToString() + "_SplInstructions where (AutoArtID='" + strAutoArtID + "' or AutoArtId='" + strJAutoID + "') and DeptCode like '%|" + Session["DeptCode"].ToString() + "|%' and (Stage='" + strStage + "' or stage='All' or Stage like substring('" + strStage + "',1,3)) order by InstDate", Session["sConnSiteDB"].ToString());
                            if (ds.Tables[0].Rows.Count == 0)
                            {
                                strRemarksFlag = "'1'";
                            }

                            string strResult = DBProc.InsertRecord("IF NOT EXISTS (Select AutoArtID from " + Session["sCustAcc"].ToString() + "_Allocation where AutoArtID='" + strAutoArtID + "' and DeptCode='" + Session["DeptCode"].ToString() + "' and Stage='" + strStage + "')Insert into " + Session["sCustAcc"].ToString() + "_Allocation (AllocatedTo, AutoArtID, AllocatedDate, Pages, DeptCode, Stage, AllocatedBy,iss, Status,ArtStageTypeID,RemStatus,AllocatedByName,AllocatedToName) values('" + strAllocatedTo + "','" + strAutoArtID + "','" + strAllocatedDate + "','" + strPages + "','" + Session["DeptCode"].ToString() + "','" + strStage + "','" + strAllocatedBy + "'," + strIss + ",'1','" + strArtStageTypeID + "', " + strRemarksFlag + ",'" + strAllocatedByName + "','" + strAllocatedToName + "')", Session["sConnSiteDB"].ToString());

                            if (strResult == "1")
                            {
                                return Json(new { dataJson = "Article " + strAutoArtID + " allocated to " + Session["EmpName"].ToString() }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json(new { dataJson = strResult + "; unable to allocate."}, JsonRequestBehavior.AllowGet);
                            }
                            
                        }
                        else
                        {
                            return Json(new { dataJson = "Article " + strAutoArtID + " already allocated to " + ds.Tables[0].Rows[0]["AllocatedToName"].ToString().Trim() }, JsonRequestBehavior.AllowGet);
                        }

                    }

                }
                else
                {
                    return Json(new { dataJson = "Failed" }, JsonRequestBehavior.AllowGet);
                }
                return Json(new { dataJson = "Success" }, JsonRequestBehavior.AllowGet);


            }
            catch (Exception)
            {
                return Json(new { dataJson = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        //// For async awit test start
        //public async Task<string> GetRemoveAllottedJobsAsync(string CheckedArticleID)
        //{
        //    await Task.Delay(3000); //Use - when you want a logical delay without blocking the current thread.  
        //    return "India";
        //}
        //public async Task<string> GetStateAsync()
        //{
        //    await Task.Delay(5000); //Use - when you want a logical delay without blocking the current thread.  
        //    return "Gujarat";
        //}
        //public async Task<ActionResult> RemoveAllottedJobsAsync(string CheckedArticleID)
        //{
        //    //Create a stopwatch for getting excution time  

        //    var watch = new Stopwatch();
        //    watch.Start();
        //    gen.WriteLog("start ... " + watch.ElapsedMilliseconds + " " + DateTime.Now.ToString());
        //    var country = GetRemoveAllottedJobsAsync(CheckedArticleID);
        //    gen.WriteLog("get remove... " + DateTime.Now.ToString());
        //    var state = GetStateAsync();
        //    gen.WriteLog("get state... " + DateTime.Now.ToString());
        //    var content = await country;
        //    gen.WriteLog("await country ... " + watch.ElapsedMilliseconds + " " + DateTime.Now.ToString());
        //    var count = await state;
        //    gen.WriteLog("await state ... " + watch.ElapsedMilliseconds + " " + DateTime.Now.ToString());
        //    gen.WriteLog("get state... " + content + " " + count + " " + DateTime.Now.ToString());
        //    watch.Stop();
        //    ViewBag.WatchMilliseconds = watch.ElapsedMilliseconds;
        //    gen.WriteLog("end ... " + watch.ElapsedMilliseconds + " " + DateTime.Now.ToString());
        //    return Json(new { dataDB = "Success " + content }, JsonRequestBehavior.AllowGet);
        //    //return View();
        //}
        //// For async awit test end




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

    }
}
