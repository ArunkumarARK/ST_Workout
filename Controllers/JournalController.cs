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
using RL = ReferenceLibrary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using System.Runtime.Serialization.Formatters.Binary;

namespace SmartTrack.Controllers
{
    public class JournalController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        DataSet dsart = new DataSet();
        string strStageinfo = string.Empty;
        string strArtcileinfo = string.Empty;
        string strProdinfo = string.Empty;
        string strProdStatus = string.Empty;
        // GET: Journal
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
            ViewBag.PageDescription = "Tell us about your Journal Details";
            return View();
        }
        [SessionExpire]
        public ActionResult JournalInfo()
        {
            try
            {
                ViewBag.PrjTabHead = "Journal Information";
                ViewBag.PageHead = "Journal Information";
                ViewBag.PageDescription = "Tell us about your Journal Details";

                //Load Team List
                List<SelectListItem> lstTeam = new List<SelectListItem>();
                DataSet dsTeam = new DataSet();
                dsTeam = DBProc.GetResultasDataSet("select TeamID,Description from JBM_CustTeamID  where CustType='" + Session["sCustAcc"].ToString() + "'", Session["sConnSiteDB"].ToString());
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

                //Load Trim Size List
                List<SelectListItem> lstcolor = new List<SelectListItem>();
                DataSet dscolor = new DataSet();
                dscolor = DBProc.GetResultasDataSet("SELECT ColorID,ColorDesc from JBM_ColorMaster", Session["sConnSiteDB"].ToString());
                if (dscolor.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dscolor.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dscolor.Tables[0].Rows[intCount]["ColorID"].ToString();
                        string strEmpName = dscolor.Tables[0].Rows[intCount]["ColorDesc"].ToString();
                        lstcolor.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Colorlist = lstcolor;

                //Load Trim Size List
                List<SelectListItem> lstTrim = new List<SelectListItem>();
                DataSet dsTrim = new DataSet();
                dsTrim = DBProc.GetResultasDataSet("Select TrimSize from JBM_TrimSize", Session["sConnSiteDB"].ToString());
                if (dsTrim.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsTrim.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsTrim.Tables[0].Rows[intCount]["TrimSize"].ToString().Trim();
                        string strEmpName = dsTrim.Tables[0].Rows[intCount]["TrimSize"].ToString().Trim();
                        lstTrim.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Trimlist = lstTrim;

                //Load Platform List
                List<SelectListItem> lstPF = new List<SelectListItem>();
                DataSet dsPF = new DataSet();
                dsPF = DBProc.GetResultasDataSet("Select PlatformID,PlatformDesc from JBM_Platform", Session["sConnSiteDB"].ToString());
                if (dsPF.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsPF.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsPF.Tables[0].Rows[intCount]["PlatformID"].ToString();
                        string strEmpName = dsPF.Tables[0].Rows[intCount]["PlatformDesc"].ToString();
                        lstPF.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.PFlist = lstPF;

                //Load Customer List
                List<SelectListItem> lstCust = new List<SelectListItem>();
                DataSet dsCust = new DataSet();
                dsCust = DBProc.GetResultasDataSet("Select CustSN,CustID from JBM_CustomerMaster  where CustType='" + Session["sCustAcc"].ToString() + "'", Session["sConnSiteDB"].ToString());
                if (dsCust.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCust.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsCust.Tables[0].Rows[intCount]["CustID"].ToString();
                        string strEmpName = dsCust.Tables[0].Rows[intCount]["CustSN"].ToString();
                        lstCust.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Custlist = lstCust;

                //Load Journal List
                List<SelectListItem> lstJournal = new List<SelectListItem>();
                DataSet dsJournal = new DataSet();
                dsJournal = DBProc.GetResultasDataSet("select JBM_AutoID,JBM_Name from JBM_info", Session["sConnSiteDB"].ToString());
                if (dsJournal.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsJournal.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsJournal.Tables[0].Rows[intCount]["JBM_AutoID"].ToString();
                        string strEmpName = dsJournal.Tables[0].Rows[intCount]["JBM_Name"].ToString();
                        lstJournal.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Journallist = lstJournal;

                //Load ExternalWorkflow List
                List<SelectListItem> lstEW = new List<SelectListItem>();
                DataSet dsEW = new DataSet();
                dsEW = DBProc.GetResultasDataSet("select distinct ops_wf from jbm_info where ops_wf is not null ", Session["sConnSiteDB"].ToString());
                //where CustID = '" + strCustID + "' AND   JBM_AutoID='" + Session["sesJobID"].ToString().Trim() + "'
                if (dsEW.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsEW.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsEW.Tables[0].Rows[intCount]["ops_wf"].ToString();
                        string strEmpName = dsEW.Tables[0].Rows[intCount]["ops_wf"].ToString();
                        lstEW.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.EWlist = lstEW;

                //Load EPT Parallel List
                List<SelectListItem> lstEP = new List<SelectListItem>();
                DataSet dsEP = new DataSet();
                dsEP = DBProc.GetResultasDataSet("select WFName from jbm_wfcode where cwp='1'", Session["sConnSiteDB"].ToString());
                if (dsEP.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsEP.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsEP.Tables[0].Rows[intCount]["WFName"].ToString();
                        string strEmpName = dsEP.Tables[0].Rows[intCount]["WFName"].ToString();
                        lstEP.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.EPlist = lstEP;

                //Load EPT Parallel List
                List<SelectListItem> lstCWPEPT = new List<SelectListItem>();
                DataSet dsCWPEPT = new DataSet();
                dsCWPEPT = DBProc.GetResultasDataSet("Select WFName from jbm_wfcode where CWP is not null order by cast(substring(WFName, 2, 5) as int)", Session["sConnSiteDB"].ToString());
                if (dsCWPEPT.Tables[0].Rows.Count > 0)
                {

                    for (int intCount = 0; intCount < dsCWPEPT.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsCWPEPT.Tables[0].Rows[intCount]["WFName"].ToString();
                        string strEmpName = dsCWPEPT.Tables[0].Rows[intCount]["WFName"].ToString();
                        lstCWPEPT.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.CWPEPTlist = lstCWPEPT;

                //Load Workflow List
                List<SelectListItem> lstWF = new List<SelectListItem>();
                DataSet dsWF = new DataSet();
                dsWF = DBProc.GetResultasDataSet("select WFName from jbm_wfcode where cwp is null", Session["sConnSiteDB"].ToString());
                if (dsWF.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsWF.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsWF.Tables[0].Rows[intCount]["WFName"].ToString();
                        string strEmpName = dsWF.Tables[0].Rows[intCount]["WFName"].ToString();
                        lstWF.Add(new SelectListItem
                        {
                            Text = strEmpName.ToString(),
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.WFlist = lstWF;



                Session["sStatus"] = null;
                if (Session["sesJobID"].ToString() != "" && Session["sesJobID"].ToString() != null)
                {
                    if (Session["sesStatus"].ToString().Trim() != "" && Session["sesStatus"].ToString().Trim() != null)
                    {
                        DataTable ds = new DataTable();
                        string strQueryFinal = @"SELECT  ji.JBM_AutoID, ji.JBM_Name, ji.JBM_IntrnlID, ji.IM_Input_JID, ji.JBM_ID, ji.JBM_Intrnl, 
ji.JBM_SubTeam, ji.Sample_wf, ji.JBM_Trimsize, ji.DirectComp, ji.ColorID, ji.Title, ji.JBM_Category, jc.CustSN, jc.Cust_Disabled, 
JBM_CustTeamID.Description as Team,ji.CustID,
JBM_IssueMan,FTPID,JBM_Disabled,JBM_mFrom,JBM_mEditor,PlatformDesc,
JBM_mReplyto,JBM_mRemainder,JBM_mBCc,JBM_mExpire,JBM_Im,JBM_Level,JBM_CeRate,JBM_Rate,JBM_Sproof,JBM_ProofType,
CASE WHEN JBM_ProofType='0' THEN 'S-Proof'  WHEN JBM_ProofType='1' THEN 'GWPS' WHEN JBM_ProofType='2' THEN 'Web Based' WHEN JBM_ProofType='3' THEN 'Smart Proof'
    WHEN JBM_ProofType='4' THEN 'Smart PDF'  ELSE '' END as JBM_ProofTypeName,Deliverables,Sheridan,
SGMLOld_wf,JBM_Platform,JBM_PeName,JBM_mPe, JBM_CustPe, WIP, JBM_mCc,JBM_ProdMan,JBM_ProdManEmail,JBM_ProdManTele,
JBM_PeRemainder,JBM_PrRemainder,JBM_EdtRemainder,JBM_eProofAttachment,HWCODE,DelyID,POD_required,
PrePrintJrnl,Composition,RevEpt_Wf,JBM_IgnoreEforms,JBM_IgnoreForms,JBM_MailFormatHTML,JBM_IgnoreEproofCmtEnable,JBM_WaterMarkLogo,
JBM_WaterMarkLogoSettings,JBM_WaterMarkNA,JBM_FormsPdfRuntime,JBM_GenEProofTpl,JBM_IgnoreArticleInProofAttachment,CWPURL,OUP_Code,
website,DOI_prefix,IM_Rate,JBM_mRemCc,Observer_MailMode,JBM_mSrEditor,JBM_SrEdtRemainder,SrEditorName,EditorRole,SrEditorRole,JBM_Editor2Name,
JBM_mEditor2_To,JBM_mEditor2_Cc,CcEditor,OnlOnly,JBM_IMMail,JBM_RevProof,ji.JBM_TeamID,JBM_CeRequired,JBM_CeRequiredOffShore,JNL_ISSNPrint,JNL_ISSNonline,
JBM_SngDbl,CASE WHEN JBM_SngDbl='0' THEN 'Single'  WHEN JBM_SngDbl='1' THEN 'Double' WHEN JBM_SngDbl='2' THEN 'Stub' WHEN JBM_SngDbl='3' THEN 'Triple'
WHEN JBM_SngDbl='4' THEN '4Column' WHEN JBM_SngDbl='5' THEN 'Multiple'  ELSE '' END as Col,PE_Approval,PAP_A,PAP_B,PE_ApprovalWait,JBM_PrinterName,PageExtent,RevisesChaserDays,FrequentIssuePerYear,
FtpLoc,Dictionary,Copyrightowner,OPS_Rev_wf,PeRole,PrName,PrEmail,PrRole,EditorName,PrAdminName,PrAdminEmail,PrAdminRole,
ObserverName,ObserverEmail,ObserverRole,ReviewOrder,SiteID,DB_Location,FreshPageStart,JBM_ColorMaster.ColorID,JBM_ColorMaster.ColorDesc,
slug,CompositePDF,PageNo,CuttingMark,PageFormat,Currency,JNL_GSM,JBM_SmallFormat,EarlyXML,onl_rel,JTF_S2008,JTF_Informa,jSAM,SAMAutEd,
Fp_wf,Rev_wf,Fin_wf,PapV_wf,PapA_wf,PapB_wf,Iss_wf,Onl_wf,SUP_wf,PRR_wf,Ver_wf ,Pes_wf,OPS_wf
FROM JBM_Info AS ji INNER JOIN
                         JBM_CustomerMaster AS jc ON ji.CustID = jc.CustID left JOIN
                        JBM_CustTeamID ON ji.JBM_TeamID = JBM_CustTeamID.TeamID  left JOIN
                          JBM_Platform ON ji.JBM_Platform = JBM_Platform.PlatformID left JOIN
						 JBM_ColorMaster ON ji.ColorID = JBM_ColorMaster.ColorID
where ji.JBM_AutoID = '" + Session["sesJobID"].ToString().Trim() + "' order by CustSN asc";
                        ds = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                        if (ds.Rows.Count > 0)
                        {
                            ViewData.Model = ds.AsEnumerable();
                            Session["sStatus"] = Session["sesStatus"].ToString();
                        }
                    }
                }
                Session["sesJobID"] = null;
                Session["sesStatus"] = null;
                return View();
            }
            catch
            {
                return View();
            }

        }
        [SessionExpire]
        public ActionResult JournalDetails()
        {
            try
            {
                //string nextStage = "";
                //nextStage = Proc_GetNextStage("C1", "S1131", "S1010");

                List<SelectListItem> lstCustomer = new List<SelectListItem>();
                DataTable ds = new DataTable();
                string strQueryFinal = @"SELECT    ji.JBM_AutoID, ji.JBM_Name, ji.JBM_IntrnlID, ji.IM_Input_JID, ji.JBM_ID, ji.JBM_Intrnl, 
ji.JBM_SubTeam, ji.Sample_wf, ji.JBM_Trimsize, ji.DirectComp, ji.ColorID, ji.Title, ji.JBM_Category, jc.CustSN, jc.Cust_Disabled, 
                         JBM_Platform.PlatformDesc, JBM_CustTeamID.Description as Team,JBM_IssueMan,ji.JBM_Trimsize,Case when ji.JBM_Disabled='0' then 'checked' else '' end as JBM_Disabled
FROM            JBM_Info AS ji INNER JOIN
                         JBM_CustomerMaster AS jc ON ji.CustID = jc.CustID INNER JOIN
                         JBM_Platform ON ji.JBM_Platform = JBM_Platform.PlatformID left JOIN
                         JBM_CustTeamID ON ji.JBM_TeamID = JBM_CustTeamID.TeamID
where  jc.CustType like '%" + Session["sCustAcc"].ToString() + "%' order by CustSN asc";
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

                ViewBag.PrjTabHead = "Journal Information";
                ViewBag.PageHead = "Journal Information";
                ViewBag.PageDescription = "Details of all the Journals";
                return View();
            }
            catch (Exception ex)
            {
                return View();
            }
        }
        [SessionExpire]
        public ActionResult ViewJournalDetails(string sJobID)
        {
            try
            {
                if (sJobID != "")
                {
                    Session["sesJobID"] = sJobID;
                    Session["sesStatus"] = "View";
                    Session["sJBM_AutoID"] = sJobID;
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetCustomerTeam(string sCustID)
        {
            try
            {
                if (sCustID != "")
                {
                    DataSet dscustlst = new DataSet();
                    dscustlst = DBProc.GetResultasDataSet(" select JBM_TeamID from JBM_CustomerMaster where CustID='" + sCustID + "'", Session["sConnSiteDB"].ToString());
                    if (dscustlst.Tables[0].Rows.Count > 0)
                    {
                        Session["CustTeam"] = dscustlst.Tables[0].Rows[0]["JBM_TeamID"].ToString();
                    }
                }
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetJournalInfo(string sJBM_AutoID)
        {
            try
            {
                if (sJBM_AutoID != "")
                {
                    Session["sesJobID"] = sJBM_AutoID;
                    Session["sesStatus"] = "template";
                }
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult EditJournalDetails(string sJobID)
        {
            try
            {
                if (sJobID != "")
                {
                    Session["sesJobID"] = sJobID;
                    Session["sesStatus"] = "Edit";
                    Session["sJBM_AutoID"] = sJobID;
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult JournalReport()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult OtherInfo()
        {
            ViewBag.PrjTabHead = "Other Information";
            ViewBag.PageHead = "Other Information";
            ViewBag.PageDescription = "Tell us about your Journal Details";
            //Load ArtTypeMT List
            List<SelectListItem> lstArtTypeMT = new List<SelectListItem>();
            DataSet dsArtTypeMT = new DataSet();
            dsArtTypeMT = DBProc.GetResultasDataSet("select ArtTypeDesc, cast(ArttypeId as varchar) as ArttypeID  from JBM_ArticleTypes where ArtTypeDesc <> '' and ArtTypeDesc <> '--'", Session["sConnSiteDB"].ToString());
            if (dsArtTypeMT.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsArtTypeMT.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsArtTypeMT.Tables[0].Rows[intCount]["ArttypeID"].ToString();
                    string strEmpName = dsArtTypeMT.Tables[0].Rows[intCount]["ArtTypeDesc"].ToString();
                    lstArtTypeMT.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.ArtTypeMTlist = lstArtTypeMT;

            //Load SubArtTypeFoot List
            List<SelectListItem> lstSubArtTypeFoot = new List<SelectListItem>();
            DataSet dsSubArtTypeFoot = new DataSet();
            dsSubArtTypeFoot = DBProc.GetResultasDataSet("select  SubArtTypeid,ArtTypeDesc from JBM_SubArticleTypes", Session["sConnSiteDB"].ToString());
            if (dsSubArtTypeFoot.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsSubArtTypeFoot.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsSubArtTypeFoot.Tables[0].Rows[intCount]["SubArtTypeid"].ToString();
                    string strEmpName = dsSubArtTypeFoot.Tables[0].Rows[intCount]["ArtTypeDesc"].ToString();
                    lstSubArtTypeFoot.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.SubArtTypeFootlist = lstSubArtTypeFoot;

            //Load ArtTypeTATFoot List
            string strJBM_AutoID = "";
            if (Session["sJBM_AutoID"] != null)
                strJBM_AutoID = Session["sJBM_AutoID"].ToString();
            List<SelectListItem> lstArtTypeTATFoot = new List<SelectListItem>();
            DataSet dsArtTypeTATFoot = new DataSet();
            dsArtTypeTATFoot = DBProc.GetResultasDataSet("select b.ArtTypeDesc, a.ArttypeID from JBM_JrnlArticleTypes a, JBM_ArticleTypes b where a.ArttypeID=b.ArtTypeID and a.jbm_AutoID='" + strJBM_AutoID + "'", Session["sConnSiteDB"].ToString());
            if (dsArtTypeTATFoot.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsArtTypeTATFoot.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsArtTypeTATFoot.Tables[0].Rows[intCount]["ArttypeID"].ToString();
                    string strEmpName = dsArtTypeTATFoot.Tables[0].Rows[intCount]["ArtTypeDesc"].ToString();
                    lstArtTypeTATFoot.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.ArtTypeTATFootlist = lstArtTypeTATFoot;

            //Load ArtTypeFoot List

            List<SelectListItem> lstArtTypeFoot = new List<SelectListItem>();
            DataSet dsArtTypeFoot = new DataSet();
            dsArtTypeFoot = DBProc.GetResultasDataSet("select b.ArtTypeDesc, a.ArttypeID from JBM_JrnlArticleTypes a, JBM_ArticleTypes b where b.ArtTypeDesc <> '' and b.ArtTypeDesc <> '--' and a.ArttypeID=b.ArtTypeID and a.jbm_AutoID='" + strJBM_AutoID + "'", Session["sConnSiteDB"].ToString());
            if (dsArtTypeFoot.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsArtTypeFoot.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsArtTypeFoot.Tables[0].Rows[intCount]["ArttypeID"].ToString();
                    string strEmpName = dsArtTypeFoot.Tables[0].Rows[intCount]["ArtTypeDesc"].ToString();
                    lstArtTypeFoot.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.ArtTypeFootlist = lstArtTypeFoot;

            //Load SubArtTypeFoot List
            List<SelectListItem> lstWorkflowFoot = new List<SelectListItem>();
            DataSet dsWorkflowFoot = new DataSet();
            dsWorkflowFoot = DBProc.GetResultasDataSet("select '' as WFName,'' as WFCode from JBM_WFCode union select WFName,WFCode from JBM_WFCode where cwp is not null", Session["sConnSiteDB"].ToString());
            if (dsWorkflowFoot.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsWorkflowFoot.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsWorkflowFoot.Tables[0].Rows[intCount]["WFName"].ToString();
                    string strEmpName = dsWorkflowFoot.Tables[0].Rows[intCount]["WFName"].ToString();
                    lstWorkflowFoot.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.WorkflowFootlist = lstWorkflowFoot;

            string strquery = @"select  distinct b.ArtTypeDesc,c.ArtTypeDesc as ArticleType, a.ArtTypeID as [ArtTypeID], a.JBM_AutoID,
 CASE
    WHEN a.Pap_Required is null THEN '0'
    ELSE a.Pap_Required
END as Pap_Required,
CASE
    WHEN a.License_Required is null THEN '0'
    ELSE a.License_Required
END as License_Required
 from JBM_JrnlArticleTypes a inner
 join JBM_ArticleTypes b on b.ArtTypeID = a.ArtTypeID left
 join JBM_SubArticleTypes c
on c.SubArtTypeID = a.Articletype where  a.JBM_AutoID = '" + strJBM_AutoID + "' ";
            DataTable dt = new DataTable();
            dt = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());
            if (dt.Rows.Count == 0)
            {
                strquery = "Select Null as [ArtTypeDesc],Null as [ArticleType], Null as [ArtTypeID], Null as [JBM_AutoId],0 as [Pap_Required],0 as [License_Required]";
                dt = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());
            }
            //Art TAT
            strquery = "select b.ArtTypeDesc, (case a.Priority when 'N' then 'Normal' when 'R' then 'Rush' when 'SR' then 'Super Rush' when 'DC' then 'Direct Comp' when 'DCR' then 'Direct Comp Rush' else a.Priority end) as [Priority], a.TAT as [TAT], a.LoginTAT as [LoginTAT], a.VendorTAT as [VendorTAT], a.ReTAT as [ReTAT], a.CeTAT as [CeTAT], a.CompTAT as [CompTAT], a.JBM_AutoID as [JBM_AutoID], b.ArtTypeID as [ArtTypeID], a.Priority as [PR], [AutoID] from JBM_TAT_Master a, JBM_ArticleTypes b where stage like '%[0-9]%' and convert(varchar(10),a.stage)=convert(varchar(10),b.ArtTypeID) and a.JBM_AutoID='" + strJBM_AutoID + "'";
            DataTable dtTAT = new DataTable();
            dtTAT = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());

            if (dtTAT.Rows.Count == 0)
            {
                strquery = "Select Null as [ArtTypeDesc], Null as [Priority], Null as [TAT], Null as [LoginTAT], Null as [VendorTAT], Null as [ReTAT], Null as [CeTAT], Null as [CompTAT], Null as [JBM_AutoId], Null as [ArtTypeID], Null as [PR], Null as [AutoID]";
                dtTAT = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());
            }

            //TAT
            SqlCommand sqlCmd = new SqlCommand();
            SqlParameter sqlParam = new SqlParameter();
            SqlDataAdapter myReader = new SqlDataAdapter();
            DataTable dtt = new DataTable();
            SqlConnection myConnection = new SqlConnection();
            myConnection = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            sqlCmd = new SqlCommand("JBM_TATMaster", myConnection);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.Parameters.Add(new SqlParameter("@JBMAutoID", SqlDbType.VarChar)).Value = strJBM_AutoID;
            myConnection.Open();
            sqlCmd.Connection = myConnection;
            myReader = new SqlDataAdapter(sqlCmd);
            myReader.Fill(dtt);
            myReader.Dispose();
            myConnection.Close();

            strquery = "Select a.ArtTypeDesc as [ArtTypeDesc],j.OPS_wf as [WFName], j.mCc as [mCc], j.mBCc as [mBCc], j.JBM_AutoId as [JBM_AutoId], a.ArtTypeId as [ArtTypeID], [AutoID],[mEditorName],[mEditor2_Name],[mPeName],[mPrName],[mEditor],[mPe],[mPr],[mEditor_cc],[mPe_cc],[mPr_cc],[mEditor2_cc],[mEditor2_To] from " + Init_Tables.gTblJBMArticleCategory + " j, " + Init_Tables.gTblArticleTypes + " a where j.ArtTypeID=a.ArtTypeID and j.JBM_AutoId='" + strJBM_AutoID + "'";
            DataTable dtart = new DataTable();
            dtart = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());
            if (dtart.Rows.Count == 0)
            {
                strquery = "Select Null as [ArtTypeDesc],Null as [WFName],Null as [mCc], Null as [mBCc], Null as [JBM_AutoId], Null as [ArtTypeID], Null as [AutoId], Null as [mEditorName],Null as [mEditor2_Name],Null as [mPeName], Null as [mPrName],Null as [mEditor],Null as [mPe],Null as [mPr],Null as [mEditor_cc],Null as [mPe_cc],Null as [mPr_cc],Null as [mEditor_cc],Null as [mPe_cc],Null as [mPr_cc],Null as [mEditor2_cc],Null as [mEditor2_To]";
                dtart = DBProc.GetResultasDataTbl(strquery, Session["sConnSiteDB"].ToString());
            }

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.Tables.Add(dtTAT);
            ds.Tables.Add(dtt);
            ds.Tables.Add(dtart);
            //Session["sJBM_AutoID"] = null;
            return View(ds);
        }
        [SessionExpire]
        public ActionResult ArticleHistory()
        {
            ViewBag.PrjTabHead = "Journal Information";
            ViewBag.PageHead = "Journal Information";
            ViewBag.PageDescription = "Tell us about your Journal Details";

            //Load Journal List
            List<SelectListItem> lstjnl = new List<SelectListItem>();
            DataSet dsjnl = new DataSet();
            dsjnl = DBProc.GetResultasDataSet("select JBM_AutoID,JBM_Name from JBM_Info", Session["sConnSiteDB"].ToString());
            if (dsjnl.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < dsjnl.Tables[0].Rows.Count; intCount++)
                {
                    string strEmpAutoID = dsjnl.Tables[0].Rows[intCount]["JBM_AutoID"].ToString();
                    string strEmpName = dsjnl.Tables[0].Rows[intCount]["JBM_Name"].ToString();
                    lstjnl.Add(new SelectListItem
                    {
                        Text = strEmpName.ToString(),
                        Value = strEmpAutoID.ToString()
                    });
                }

            }

            ViewBag.Userlist = lstjnl;

            List<SelectListItem> lstTeam = new List<SelectListItem>();
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet("SELECT   Seq_ID, Tables_Name, Columns_Name, Data_Type, Display_Name, InputCtrl_Type, ItemsFromQuery, Visibility, Field_Desc, Modified_By, Modified_Date FROM  JBM_FieldMapDetails where Visibility='Yes'", Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                // Creating single column into three column view
                DataSet dsNew = new DataSet();

                if (ds.Tables[0].Rows.Count > 0)
                {
                    int ColCnt = 1;

                    DataTable dt = new DataTable("MyTable");
                    dt.Columns.Add(new DataColumn("Col1Column_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col1", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col1Display_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col1Tables_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col1Input", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col2Column_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col2", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col2Display_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col2Tables_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col2Input", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col3Column_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col3", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col3Display_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col3Tables_Name", typeof(string)));
                    dt.Columns.Add(new DataColumn("Col3Input", typeof(string)));

                    DataRow dr = dt.NewRow();
                    string strCol1Val = "";
                    string strCol1Col = "";
                    string strCol1Dis = "";
                    string strCol1Tab = "";
                    string strCol1Inp = "";
                    string strCol2Val = "";
                    string strCol2Col = "";
                    string strCol2Dis = "";
                    string strCol2Tab = "";
                    string strCol2Inp = "";
                    string strCol3Val = "";
                    string strCol3Col = "";
                    string strCol3Dis = "";
                    string strCol3Tab = "";
                    string strCol3Inp = "";
                    int newRow = 0;
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {


                        if (ColCnt == 1)
                        {
                            strCol1Val = ds.Tables[0].Rows[intCount]["Seq_ID"].ToString();
                            strCol1Col = ds.Tables[0].Rows[intCount]["Columns_Name"].ToString();
                            strCol1Dis = ds.Tables[0].Rows[intCount]["Display_Name"].ToString();
                            strCol1Tab = ds.Tables[0].Rows[intCount]["Tables_Name"].ToString();
                            strCol1Inp = ds.Tables[0].Rows[intCount]["InputCtrl_Type"].ToString();
                            ColCnt = 2;
                        }
                        else if (ColCnt == 2)
                        {
                            strCol2Val = ds.Tables[0].Rows[intCount]["Seq_ID"].ToString();
                            strCol2Col = ds.Tables[0].Rows[intCount]["Columns_Name"].ToString();
                            strCol2Dis = ds.Tables[0].Rows[intCount]["Display_Name"].ToString();
                            strCol2Tab = ds.Tables[0].Rows[intCount]["Tables_Name"].ToString();
                            strCol2Inp = ds.Tables[0].Rows[intCount]["InputCtrl_Type"].ToString();
                            ColCnt = 3;

                        }
                        else if (ColCnt == 3)
                        {
                            strCol3Val = ds.Tables[0].Rows[intCount]["Seq_ID"].ToString();
                            strCol3Col = ds.Tables[0].Rows[intCount]["Columns_Name"].ToString();
                            strCol3Dis = ds.Tables[0].Rows[intCount]["Display_Name"].ToString();
                            strCol3Tab = ds.Tables[0].Rows[intCount]["Tables_Name"].ToString();
                            strCol3Inp = ds.Tables[0].Rows[intCount]["InputCtrl_Type"].ToString();
                            newRow = 1;
                        }


                        if (ColCnt == 2 && intCount == ds.Tables[0].Rows.Count - 1)
                        {
                            newRow = 1;
                        }
                        else if (ColCnt == 3 && intCount == ds.Tables[0].Rows.Count - 1)
                        {
                            newRow = 1;
                        }

                        if (newRow == 1)
                        {
                            dr = dt.NewRow();
                            dr["Col1Column_Name"] = strCol1Col;
                            dr["Col1"] = strCol1Val;
                            dr["Col1Display_Name"] = strCol1Dis;
                            dr["Col1Tables_Name"] = strCol1Tab;
                            dr["Col1Input"] = strCol1Inp;
                            dr["Col2Column_Name"] = strCol2Col;
                            dr["Col2"] = strCol2Val;
                            dr["Col2Display_Name"] = strCol2Dis;
                            dr["Col2Tables_Name"] = strCol2Tab;
                            dr["Col2Input"] = strCol2Inp;
                            dr["Col3Column_Name"] = strCol3Col;
                            dr["Col3"] = strCol3Val;
                            dr["Col3Display_Name"] = strCol3Dis;
                            dr["Col3Tables_Name"] = strCol3Tab;
                            dr["Col3Input"] = strCol3Inp;

                            dt.Rows.Add(dr);


                            strCol1Val = "";
                            strCol2Val = "";
                            strCol3Val = "";
                            strCol1Col = "";
                            strCol2Col = "";
                            strCol3Col = "";
                            strCol1Dis = "";
                            strCol2Dis = "";
                            strCol3Dis = "";
                            strCol1Tab = "";
                            strCol2Tab = "";
                            strCol3Tab = "";
                            strCol1Inp = "";
                            strCol2Inp = "";
                            strCol3Inp = "";
                            ColCnt = 1;
                            newRow = 0;
                        }


                    }
                    dsNew.Tables.Add(dt);
                }
                ViewData.Model = dsNew.Tables[0].AsEnumerable();
            }

            return View();
        }
        [SessionExpire]
        public ActionResult GetColumnDetails(List<Chklst> Columnlst)
        {
            try
            {
                int count = Columnlst.Count;
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                for (int i = 0; i < count; i++)
                {
                    string sid = Columnlst[i].Chkid.ToString().Trim();
                    if (sid.Trim() != "")
                    {
                        string strQueryFinal = "SELECT Tables_Name, Columns_Name, Data_Type, Display_Name, InputCtrl_Type, ItemsFromQuery FROM JBM_FieldMapDetails where Visibility = 'Yes' and Seq_ID = '" + sid + "'";

                        dt = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                        string UniqueName = dt.Rows[0]["Columns_Name"].ToString().Trim();
                        string InputCtrl_Type = dt.Rows[0]["InputCtrl_Type"].ToString().Trim();
                        string Display_Name = dt.Rows[0]["Display_Name"].ToString().Trim();
                        if (dt.Rows.Count > 0)
                        {
                            if (InputCtrl_Type == "Dropdown")
                            {
                                strQueryFinal = dt.Rows[0]["ItemsFromQuery"].ToString().Trim();
                                dt = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                                dt.Columns.Add("Unique_Name");
                                dt.Columns.Add("Ctrl_Type");
                                dt.Columns.Add("Count");
                                dt.Columns.Add("Display_Name");
                                for (int j = 0; j < dt.Rows.Count; j++)
                                {
                                    DataRow dr = dt.Rows[j];
                                    dr[2] = UniqueName;
                                    dr[3] = InputCtrl_Type;
                                    dr[5] = Display_Name;
                                }
                                if (dt.Rows.Count == 1)
                                {
                                    DataRow dr1 = dt.Rows[0];
                                    dr1[4] = "2";
                                }
                                else
                                {
                                    DataRow dr1 = dt.Rows[0];
                                    dr1[4] = "0";
                                    DataRow dr2 = dt.Rows[dt.Rows.Count - 1];
                                    dr2[4] = "1";
                                }
                                ds.Tables.Add(dt);
                            }
                            else if (InputCtrl_Type == "Checkbox")
                            {
                                strQueryFinal = dt.Rows[0]["ItemsFromQuery"].ToString().Trim();
                                dt = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                                DataTable dtnew = new DataTable();
                                dtnew.Clear();
                                dtnew.Columns.Add("ID");
                                dtnew.Columns.Add("Name");
                                for (int j = 0; j < dt.Columns.Count; j++)
                                {
                                    DataRow dr = dtnew.NewRow();
                                    dr["ID"] = "Check" + j;
                                    dr["Name"] = dt.Rows[0][j].ToString();
                                    dtnew.Rows.Add(dr);
                                }
                                dtnew.Columns.Add("Unique_Name");
                                dtnew.Columns.Add("Ctrl_Type");
                                dtnew.Columns.Add("Count");
                                dtnew.Columns.Add("Display_Name");
                                for (int j = 0; j < dtnew.Rows.Count; j++)
                                {
                                    DataRow dr = dtnew.Rows[j];
                                    dr[2] = UniqueName;
                                    dr[3] = InputCtrl_Type;
                                    dr[5] = Display_Name;
                                }
                                DataRow dr1 = dtnew.Rows[0];
                                dr1[4] = "0";
                                DataRow dr2 = dtnew.Rows[dtnew.Rows.Count - 1];
                                dr2[4] = "1";
                                ds.Tables.Add(dtnew);
                            }
                            else if (InputCtrl_Type == "Option")
                            {
                                strQueryFinal = dt.Rows[0]["ItemsFromQuery"].ToString().Trim();
                                dt = DBProc.GetResultasDataTbl(strQueryFinal, Session["sConnSiteDB"].ToString());
                                DataTable dtnew = new DataTable();
                                dtnew.Clear();
                                dtnew.Columns.Add("ID");
                                dtnew.Columns.Add("Name");
                                for (int j = 0; j < dt.Columns.Count; j++)
                                {
                                    DataRow dr = dtnew.NewRow();
                                    dr["ID"] = "Option" + j;
                                    dr["Name"] = dt.Rows[0][j].ToString();
                                    dtnew.Rows.Add(dr);
                                }
                                dtnew.Columns.Add("Unique_Name");
                                dtnew.Columns.Add("Ctrl_Type");
                                dtnew.Columns.Add("Count");
                                dtnew.Columns.Add("Display_Name");
                                for (int j = 0; j < dtnew.Rows.Count; j++)
                                {
                                    DataRow dr = dtnew.Rows[j];
                                    dr[2] = UniqueName;
                                    dr[3] = InputCtrl_Type;
                                    dr[5] = Display_Name;
                                }
                                DataRow dr1 = dtnew.Rows[0];
                                dr1[4] = "0";
                                DataRow dr2 = dtnew.Rows[dtnew.Rows.Count - 1];
                                dr2[4] = "1";
                                ds.Tables.Add(dtnew);
                            }
                            else if (InputCtrl_Type == "Textbox")
                            {
                                dt = new DataTable();
                                dt.Columns.Add("ID");
                                dt.Columns.Add("Name");
                                dt.Columns.Add("Unique_Name");
                                dt.Columns.Add("Ctrl_Type");
                                dt.Columns.Add("Count");
                                dt.Columns.Add("Display_Name");

                                DataRow dr = dt.NewRow();
                                dr[0] = UniqueName;
                                dr[1] = UniqueName;
                                dr[2] = UniqueName;
                                dr[3] = InputCtrl_Type;
                                dr[4] = "0";
                                dr[5] = Display_Name;
                                dt.Rows.Add(dr);
                                ds.Tables.Add(dt);
                            }
                            else if (InputCtrl_Type == "DATETIME")
                            {
                                dt = new DataTable();
                                dt.Columns.Add("ID");
                                dt.Columns.Add("Name");
                                dt.Columns.Add("Unique_Name");
                                dt.Columns.Add("Ctrl_Type");
                                dt.Columns.Add("Count");
                                dt.Columns.Add("Display_Name");

                                DataRow dr = dt.NewRow();
                                dr[0] = UniqueName;
                                dr[1] = UniqueName;
                                dr[2] = UniqueName;
                                dr[3] = InputCtrl_Type;
                                dr[4] = "0";
                                dr[5] = Display_Name;
                                dt.Rows.Add(dr);
                                ds.Tables.Add(dt);
                            }
                        }
                    }
                    else
                    {
                        return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
                    }

                }
                int tblcount = ds.Tables.Count;
                DataTable dtnew1 = new DataTable();
                dtnew1.Clear();
                dtnew1.Columns.Add("ID");
                dtnew1.Columns.Add("Name");
                dtnew1.Columns.Add("Unique_Name");
                dtnew1.Columns.Add("Ctrl_Type");
                dtnew1.Columns.Add("Count");
                dtnew1.Columns.Add("Display_Name");
                for (int i = 0; i < tblcount; i++)
                {
                    for (int j = 0; j < ds.Tables[i].Rows.Count; j++)
                    {
                        DataRow dr = dtnew1.NewRow();
                        dr["ID"] = ds.Tables[i].Rows[j][0].ToString();
                        dr["Name"] = ds.Tables[i].Rows[j][1].ToString();
                        dr["Unique_Name"] = ds.Tables[i].Rows[j][2].ToString();
                        dr["Ctrl_Type"] = ds.Tables[i].Rows[j][3].ToString();
                        dr["Count"] = ds.Tables[i].Rows[j][4].ToString();
                        dr["Display_Name"] = ds.Tables[i].Rows[j][5].ToString();
                        dtnew1.Rows.Add(dr);
                    }
                }

                var JSONString = from a in dtnew1.AsEnumerable()
                                 select new[] {a[0].ToString(),a[1].ToString(),a[2].ToString(),a[3].ToString(),a[4].ToString(),a[5].ToString()
                                };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);


            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult UpdateColumnDetails(List<Chklst1> Columnlst)
        {
            try
            {
                int count = Columnlst.Count;
                string Jid = Columnlst[0].Empid.ToString().Trim();
                string chklst = "";
                string strResult = "UPDATE JBM_Info SET ";
                for (int i = 0; i < count; i++)
                {
                    if (chklst != "")
                        chklst = chklst + "," + Columnlst[i].Chkid.ToString().Trim();
                    else
                        chklst = Columnlst[i].Chkid.ToString().Trim();
                    strResult += " " + Columnlst[i].Chkid.ToString().Trim() + "='" + Columnlst[i].ChkValue.ToString().Trim() + "',";
                }
                if (Jid.Trim() != "")
                {
                    //add xml
                    XmlDocument xmlFileDoc = new XmlDocument();
                    xmlFileDoc.LoadXml("<Fields></Fields>");
                    XmlElement FieldsList = xmlFileDoc.CreateElement("FieldsList");
                    FieldsList.InnerText = chklst;
                    xmlFileDoc.DocumentElement.AppendChild(FieldsList);
                    string XMLDetails = xmlFileDoc.InnerXml;
                    strResult = DBProc.GetResultasString(strResult + "JBM_FieldInfo='" + XMLDetails + "' WHERE JBM_AutoID='" + Jid + "'", Session["sConnSiteDB"].ToString());
                }
                else
                {
                    return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
                }


                //var JSONString = from a in dtnew1.AsEnumerable()
                //                 select new[] {a[0].ToString(),a[1].ToString(),a[2].ToString(),a[3].ToString(),a[4].ToString(),a[5].ToString()
                //                };
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);


            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult GetJournalFieldlstData(List<Chklst> Columnlst)
        {
            try
            {
                List<ChkJournal> Customeritems = new List<ChkJournal>();
                int count = Columnlst.Count;
                string Jid = Columnlst[0].Empid.ToString().Trim();
                string chklst = "";
                for (int i = 0; i < count; i++)
                {
                    if (chklst != "")
                        chklst = chklst + "," + Columnlst[i].Chkid.ToString().Trim();
                    else
                        chklst = Columnlst[i].Chkid.ToString().Trim();
                    string FieldDataVal = "";
                    string FieldDataInput = "";
                    DataSet fielddata = new DataSet();
                    fielddata = DBProc.GetResultasDataSet(" select " + Columnlst[i].Chkid.ToString().Trim() + " from JBM_Info where JBM_AutoID='" + Jid + "'", Session["sConnSiteDB"].ToString());
                    if (fielddata.Tables[0].Rows.Count > 0)
                    {
                        FieldDataVal = fielddata.Tables[0].Rows[0][Columnlst[i].Chkid.ToString().Trim()].ToString();
                    }
                    DataSet fielddataInput = new DataSet();
                    fielddataInput = DBProc.GetResultasDataSet("select InputCtrl_Type  from JBM_FieldMapDetails where Columns_Name='" + Columnlst[i].Chkid.ToString().Trim() + "'", Session["sConnSiteDB"].ToString());
                    if (fielddataInput.Tables[0].Rows.Count > 0)
                    {
                        FieldDataInput = fielddataInput.Tables[0].Rows[0]["InputCtrl_Type"].ToString();
                    }
                    Customeritems.Add(new ChkJournal
                    {
                        Text = FieldDataVal,
                        Value = Columnlst[i].Chkid.ToString().Trim(),
                        Input = FieldDataInput
                    });
                }


                return Json(Customeritems, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult GetJournalFieldlst(string Jid)
        {
            try
            {
                List<ChkJournal> Customeritems = new List<ChkJournal>();
                DataSet dscustlst = new DataSet();
                dscustlst = DBProc.GetResultasDataSet(" select JBM_FieldInfo from JBM_Info where JBM_AutoID='" + Jid + "'", Session["sConnSiteDB"].ToString());
                if (dscustlst.Tables[0].Rows.Count > 0)
                {
                    XmlNodeList xmlnode;
                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.LoadXml(dscustlst.Tables[0].Rows[0]["JBM_FieldInfo"].ToString());
                    xmlnode = xmldoc.GetElementsByTagName("FieldsList");
                    string fields = xmlnode[0].InnerXml.ToString();
                    string[] fieldslst = fields.Split(',');
                    int count = fieldslst.Count();
                    for (int i = 0; i < count; i++)
                    {
                        string FieldDataVal = "";
                        string FieldDataInput = "";
                        DataSet fielddata = new DataSet();
                        fielddata = DBProc.GetResultasDataSet(" select " + fieldslst[i].ToString() + " from JBM_Info where JBM_AutoID='" + Jid + "'", Session["sConnSiteDB"].ToString());
                        if (fielddata.Tables[0].Rows.Count > 0)
                        {
                            FieldDataVal = fielddata.Tables[0].Rows[0][fieldslst[i].ToString()].ToString();
                        }
                        DataSet fielddataInput = new DataSet();
                        fielddataInput = DBProc.GetResultasDataSet("select InputCtrl_Type  from JBM_FieldMapDetails where Columns_Name='" + fieldslst[i].ToString() + "'", Session["sConnSiteDB"].ToString());
                        if (fielddataInput.Tables[0].Rows.Count > 0)
                        {
                            FieldDataInput = fielddataInput.Tables[0].Rows[0]["InputCtrl_Type"].ToString();
                        }
                        Customeritems.Add(new ChkJournal
                        {
                            Text = FieldDataVal,
                            Value = fieldslst[i].ToString(),
                            Input = FieldDataInput
                        });
                    }
                }
                return Json(Customeritems, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult Split_Function_WF(string WFName)
        {
            try
            {
                string data = "";
                DataSet dscustlst = new DataSet();
                dscustlst = DBProc.GetResultasDataSet("select dbo.Split_Function_WF('" + WFName + "')", Session["sConnSiteDB"].ToString());
                if (dscustlst.Tables[0].Rows.Count > 0)
                {
                    data = dscustlst.Tables[0].Rows[0][0].ToString();
                }
                return Json(new { dataSch = data }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
       
        [SessionExpire]
        [HttpPost]
        public ActionResult AddNewJournalInfo(FormCollection formCollection)
        {
            try
            {
                string strJBMAutoMaxID = "";
                string sTeam = "";
                formCollection["PERole"] = formCollection["PERole"].Replace(",", "|") + "|";
                formCollection["PRRole"] = formCollection["PRRole"].Replace(",", "|") + "|";
                formCollection["EditorRole"] = formCollection["EditorRole"].Replace(",", "|") + "|";
                formCollection["SeniorEditorRole"] = formCollection["SeniorEditorRole"].Replace(",", "|") + "|";
                formCollection["ProofAdminRole"] = formCollection["ProofAdminRole"].Replace(",", "|") + "|";
                formCollection["ObserverRole"] = formCollection["ObserverRole"].Replace(",", "|") + "|";
                formCollection["CWPURL"] = formCollection["CWPURL"].ToString() == "undefined" ? "" : formCollection["CWPURL"];
                formCollection["DView"] = formCollection["DView"].ToString() == "undefined" ? "" : formCollection["DView"];
                string strAutoIdFilter = Session["sCustAcc"].ToString();
                if (formCollection["Team"] == "")
                    sTeam = Session["CustTeam"].ToString();
                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl("select max(JBM_AutoID) as JBM_AutoID from  jbm_info where JBM_AutoID like '%" + strAutoIdFilter + "%'", Session["sConnSiteDB"].ToString());

                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["JBM_AutoID"].ToString().Trim() != "")
                    {
                        strJBMAutoMaxID = dt.Rows[0]["JBM_AutoID"].ToString();
                        var result = Regex.Match(strJBMAutoMaxID, @"\d+$").Value;
                        strJBMAutoMaxID = (Convert.ToInt32(result) + 1).ToString();

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
                }
                else
                {
                    strJBMAutoMaxID = strAutoIdFilter + "001";
                }

                string strAutoSeq = "";
                string strJBMIntrnlIdTemp = "";
                dt = DBProc.GetResultasDataTbl("SELECT COALESCE(MAX(AutoSeq), 0) as AutoSeq FROM jbm_info where custId='" + formCollection["CustID"] + "'", Session["sConnSiteDB"].ToString());
                if (dt.Rows.Count > 0)
                {
                    strAutoSeq = dt.Rows[0]["AutoSeq"].ToString();
                    strAutoSeq = (Convert.ToInt32(strAutoSeq) + 1).ToString();
                }
                else
                {
                    strAutoSeq = "1";
                }
                //
                string strFieldformsPDfRuntime = formCollection["FormPDFRuntime"].ToString().Trim() == "0" ? "null" : "'" + formCollection["FormPDFRuntime"].ToString().Trim() + "'";
                string strGenEProofTpl = formCollection["GenEProof"].ToString().Trim() != "" ? "'" + formCollection["GenEProof"].ToString().Trim() + "'" : "null";


                string SiteID = GlobalVariables.strSiteID;
                string DB_Location = GlobalVariables.strSiteCode;
                string strQuery = @"Insert into jbm_info(AutoSeq, JBM_Intrnl, JBM_AutoID, JBM_ID, JBM_Name, Title, JBM_IssueMan,FTPID,JBM_Disabled,JBM_mFrom,JBM_mEditor,
JBM_mReplyto,JBM_mRemainder,JBM_mBCc,JBM_mExpire,JBM_Im,JBM_Level,JBM_CeRate,JBM_Rate,JBM_Sproof,JBM_ProofType,Deliverables,Sheridan,
SGMLOld_wf,JBM_Platform,JBM_Trimsize,CustID,JBM_PeName,JBM_mPe, JBM_CustPe, WIP, JBM_mCc,JBM_ProdMan,JBM_ProdManEmail,JBM_ProdManTele,
JBM_PeRemainder,JBM_PrRemainder,JBM_EdtRemainder,JBM_eProofAttachment,IM_INPUT_JID,HWCODE,DelyID,POD_required,
PrePrintJrnl,Composition,RevEpt_Wf,JBM_IgnoreEforms,JBM_IgnoreForms,JBM_MailFormatHTML,JBM_IgnoreEproofCmtEnable,JBM_WaterMarkLogo,
JBM_WaterMarkLogoSettings,JBM_WaterMarkNA,JBM_FormsPdfRuntime,JBM_GenEProofTpl,JBM_IgnoreArticleInProofAttachment,CWPURL,OUP_Code,
website,DOI_prefix,IM_Rate,JBM_mRemCc,Observer_MailMode,JBM_mSrEditor,JBM_SrEdtRemainder,SrEditorName,EditorRole,SrEditorRole,JBM_Editor2Name,
JBM_mEditor2_To,JBM_mEditor2_Cc,CcEditor,OnlOnly,JBM_IMMail,JBM_RevProof,JBM_TeamID,JBM_CeRequired,JBM_CeRequiredOffShore,JNL_ISSNPrint,JNL_ISSNonline,
JBM_SngDbl,PE_Approval,PAP_A,PAP_B,PE_ApprovalWait,JBM_PrinterName,PageExtent,RevisesChaserDays,FrequentIssuePerYear,
FtpLoc,Dictionary,Copyrightowner,OPS_Rev_wf,PeRole,PrName,PrEmail,PrRole,EditorName,PrAdminName,PrAdminEmail,PrAdminRole,
ObserverName,ObserverEmail,ObserverRole,ReviewOrder,SiteID,DB_Location,FreshPageStart,ColorID,slug,PageNo,CompositePDF,CuttingMark,
PageFormat,Currency,JNL_GSM,JBM_SmallFormat,EarlyXML,onl_rel,JTF_S2008,JTF_Informa,jSAM,SAMAutEd,
Fp_wf,Rev_wf,Fin_wf,PapV_wf,PapA_wf,PapB_wf,Iss_wf,Onl_wf,SUP_wf,PRR_wf,Ver_wf ,Pes_wf,OPS_wf) 
values ('" + strAutoSeq + "','" + formCollection["InternalID"] + "','" + strJBMAutoMaxID + "','" + formCollection["InternalID"] + "','" + formCollection["JournalName"] + "','" + formCollection["Title"] + "','"
+ formCollection["IssueManager"] + "','" + formCollection["FTPID"] + "','" + formCollection["Disabled"] + "','" + formCollection["FromEmail"] + "','" + formCollection["EEmail"] + "','" + formCollection["ReplyEmail"] + "','" + formCollection["Authorchaserdays"] + "'," +
"'" + formCollection["BCC"] + "','" + formCollection["Proofexpire"] + "','" + formCollection["IMJournal"] + "','" + formCollection["LevelID"] + "','" + formCollection["CERate"] + "','" + formCollection["JournalRate"] + "','" + formCollection["EProofReq"] + "'" +
",'" + formCollection["ProofType"] + "','" + formCollection["DView"] + "','" + formCollection["rSheridan"] + "','" + formCollection["rSGMLOld"] + "','" + formCollection["PlatformID"] + "','" + formCollection["TrimSize"] + "','" + formCollection["CustID"] + "','" + formCollection["PEName"] + "'" +
",'" + formCollection["PEEmail"] + "','" + formCollection["CustomerPEEmail"] + "','1','" + formCollection["CC"] + "','" + formCollection["PMName"] + "','" + formCollection["PMEmail"] + "','" + formCollection["PMTele"] + "','" + formCollection["PERoleChaserDay"]
+ "','" + formCollection["PRRoleChaserDay"] + "','" + formCollection["ERoleChaserDay"] + "','" + formCollection["MailAttachment"] + "','" + formCollection["IMInputJID"] + "','" + formCollection["HWCode"] + "','" + formCollection["DelyID"]
+ "','" + formCollection["PrintOnDelivery"] + "','" + formCollection["PreprintJournal"] + "','" + formCollection["Composition"] + "','" + formCollection["RevisesEPTWF"] + "','" + formCollection["IgnoreEForms"] + "','" +
formCollection["IgnoreForms"] + "','" + formCollection["MailFormatHtml"] + "','" + formCollection["IgnoreEproofCmtEnable"] + "','" + formCollection["WaterMarkLogo"] + "','" + formCollection["WMLogoSettings"] + "','" +
formCollection["WaterMarkNA"] + "'," + strFieldformsPDfRuntime + "," + strGenEProofTpl + ",'" + formCollection["IgnoreArticleInProofAttch"] + "','" + formCollection["CWPURL"] + "','" +
formCollection["OUPCode"] + "','" + formCollection["JournalWebsite"] + "','" + formCollection["DOIPrefix"] + "','" + formCollection["IMRate"] + "','" + formCollection["ReminderCC"] + "','" + formCollection["CWPBccMail"] + "','" + formCollection["SEEmail"] + "','" +
formCollection["SERoleChaserDay"] + "','" + formCollection["SEName"] + "','" + formCollection["EditorRole"] + "','" + formCollection["SeniorEditorRole"] + "','" + formCollection["Editor2Name"] + "','" + formCollection["Editor2Email"] + "'" +
",'" + formCollection["Editor2CC"] + "','" + formCollection["EditorCC"] + "','" + formCollection["OnlineOnly"] + "','" + formCollection["IssueManagerEmail"] + "','" + formCollection["ReviseseProofReq"] + "','" + sTeam + "','" + formCollection["CopyEditing"] + "','" + formCollection["CopyEditingoffshore"] + "','" + formCollection["PrintISSN"] + "'" +
",'" + formCollection["OnlineISSN"] + "','" + formCollection["Column"] + "','" + formCollection["PEApproval"] + "','" + formCollection["ModelAPAP"] + "','" + formCollection["ModelBPAP"] + "','" + formCollection["WaitingPE"] + "','" + formCollection["PrinterName"] + "','" + formCollection["PageExtent"] + "'," +
"'" + formCollection["RevisesAuthorChasing"] + "','" + formCollection["Frequency"] + "','" + formCollection["USUKJournals"] + "','" + formCollection["Dictionary"] + "','" + formCollection["CopyRightOwner"] + "','" + formCollection["RevisesSPWorkflow"] + "'," +
"'" + formCollection["PERole"] + "','" + formCollection["PRName"] + "','" + formCollection["PREmail"] + "','" + formCollection["PRRole"] + "','" + formCollection["EName"] + "','" + formCollection["PAName"] + "','" + formCollection["PAEmail"] + "','" + formCollection["ProofAdminRole"] + "','" + formCollection["OName"] + "'," +
"'" + formCollection["OEmail"] + "','" + formCollection["ObserverRole"] + "','" + formCollection["EPTParallel"] + "','" + SiteID + "','" + DB_Location + "','" + formCollection["FreshPageStart"] + "','" + formCollection["Color"] + "','" + formCollection["Slug"] + "','" + formCollection["PageNo"] + "','" + formCollection["CompositePDF"] + "','" + formCollection["CuttingMark"] + "'," +
"'" + formCollection["Pageformat"] + "','" + formCollection["Currency"] + "','" + formCollection["GSM"] + "','" + formCollection["LargeFormat"] + "','" + formCollection["EarlyXML"] + "','" + formCollection["PrvParallel"] + "','" + formCollection["JTF_S2008"] + "','" + formCollection["Informa"] + "','" + formCollection["SAM"] + "','" + formCollection["SAMAutEd"] + "'" +
",'" + formCollection["Fp_wf"] + "','" + formCollection["Rev_wf"] + "','" + formCollection["Fin_wf"] + "','" + formCollection["PapV_wf"] + "','" + formCollection["PapA_wf"] + "','" + formCollection["PapB_wf"] + "','" + formCollection["Iss_wf"] + "','" + formCollection["Onl_wf"] + "'" +
",'" + formCollection["SUP_wf"] + "','" + formCollection["PRR_wf"] + "','" + formCollection["Ver_wf"] + "','" + formCollection["Pes_wf"] + "','" + formCollection["OPS_wf"] + "')";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                //strQuery = DBProc.GetResultasString("Insert into BK_ProjectDetails(JBM_AutoID) Values (" + strJBMAutoMaxID + ")", Session["sConnSiteDB"].ToString());

                //Folder Creation 
                DataTable dtrootid = new DataTable();
                string CustSN = DBProc.GetResultasString("select CustSN from JBM_CustomerMaster where CustID='" + formCollection["CustID"] + "'", Session["sConnSiteDB"].ToString());
                dtrootid = DBProc.GetResultasDataTbl("select * from JBM_RootDirectory where RootPath like '%" + CustSN + "%'", Session["sConnSiteDB"].ToString());
                for (int i = 0; i < dtrootid.Rows.Count; i++)
                {
                    string Rootpath = @"" + dtrootid.Rows[i]["RootPath"].ToString().Trim() + formCollection["InternalID"];

                    if (!Directory.Exists(Rootpath))
                    {
                        Directory.CreateDirectory(Rootpath);
                    }
                }
                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalEproof(FormCollection formCollection)
        {
            try
            {

                formCollection["PERole"] = formCollection["PERole"].ToString() == "" ? "" : formCollection["PERole"].Replace(",", "|") + "|";
                formCollection["PRRole"] = formCollection["PRRole"].ToString() == "" ? "" : formCollection["PRRole"].Replace(",", "|") + "|";
                formCollection["EditorRole"] = formCollection["EditorRole"].ToString() == "" ? "" : formCollection["EditorRole"].Replace(",", "|") + "|";
                formCollection["SeniorEditorRole"] = formCollection["SeniorEditorRole"].ToString() == "" ? "" : formCollection["SeniorEditorRole"].Replace(",", "|") + "|";
                formCollection["ProofAdminRole"] = formCollection["ProofAdminRole"].ToString() == "" ? "" : formCollection["ProofAdminRole"].Replace(",", "|") + "|";
                formCollection["ObserverRole"] = formCollection["ObserverRole"].ToString() == "" ? "" : formCollection["ObserverRole"].Replace(",", "|") + "|";
                string strAutoIdFilter = Session["sCustAcc"].ToString();
                formCollection["CWPURL"] = formCollection["CWPURL"].ToString() == "undefined" ? "" : formCollection["CWPURL"];

                string strFieldformsPDfRuntime = formCollection["FormPDFRuntime"].ToString().Trim() == "0" ? "null" : "'" + formCollection["FormPDFRuntime"].ToString().Trim() + "'";
                string strGenEProofTpl = formCollection["GenEProof"].ToString().Trim() != "" ? "'" + formCollection["GenEProof"].ToString().Trim() + "'" : "null";

                string strQuery = @"update jbm_info set OPS_Rev_wf='" + formCollection["RevisesSPWorkflow"] + "', CWPURL='" + formCollection["CWPURL"] + "',JBM_mFrom='" + formCollection["FromEmail"] + "',JBM_mEditor='" + formCollection["EEmail"] + "',JBM_mReplyto='" + formCollection["ReplyEmail"] + "',JBM_mRemainder='" + formCollection["Authorchaserdays"] + "'," +
"JBM_mBCc='" + formCollection["BCC"] + "',JBM_mExpire='" + formCollection["Proofexpire"] + "',JBM_Sproof='" + formCollection["EProofReq"] + "'" +
",JBM_ProofType='" + formCollection["ProofType"] + "',JBM_PeName='" + formCollection["PEName"] + "'" +
",JBM_mPe='" + formCollection["PEEmail"] + "',JBM_CustPe='" + formCollection["CustomerPEEmail"] + "',JBM_mCc='" + formCollection["CC"] + "',JBM_ProdMan='" + formCollection["PMName"] + "',JBM_ProdManEmail='" + formCollection["PMEmail"] + "',JBM_ProdManTele='" + formCollection["PMTele"] + "',JBM_PeRemainder='" + formCollection["PERoleChaserDay"]
+ "',JBM_PrRemainder='" + formCollection["PRRoleChaserDay"] + "',JBM_EdtRemainder='" + formCollection["ERoleChaserDay"] + "',JBM_eProofAttachment='" + formCollection["MailAttachment"] + "',JBM_IgnoreEforms='" + formCollection["IgnoreEForms"] + "',JBM_IgnoreForms='" +
formCollection["IgnoreForms"] + "',JBM_MailFormatHTML='" + formCollection["MailFormatHtml"] + "',JBM_IgnoreEproofCmtEnable='" + formCollection["IgnoreEproofCmtEnable"] + "',JBM_WaterMarkLogo='" + formCollection["WaterMarkLogo"] + "',JBM_WaterMarkLogoSettings='" + formCollection["WMLogoSettings"] + "',JBM_WaterMarkNA='" +
formCollection["WaterMarkNA"] + "',JBM_FormsPdfRuntime=" + strFieldformsPDfRuntime + ",JBM_GenEProofTpl=" + strGenEProofTpl + ",JBM_IgnoreArticleInProofAttachment='" + formCollection["IgnoreArticleInProofAttch"] + "',JBM_mRemCc='" + formCollection["ReminderCC"] + "',Observer_MailMode='" + formCollection["CWPBccMail"] + "',JBM_mSrEditor='" + formCollection["SEEmail"] + "',JBM_SrEdtRemainder='" +
formCollection["SERoleChaserDay"] + "',SrEditorName='" + formCollection["SEName"] + "',EditorRole='" + formCollection["EditorRole"] + "',SrEditorRole='" + formCollection["SeniorEditorRole"] + "',JBM_Editor2Name='" + formCollection["Editor2Name"] + "',JBM_mEditor2_To='" + formCollection["Editor2Email"] + "'" +
",JBM_mEditor2_Cc='" + formCollection["Editor2CC"] + "',CcEditor='" + formCollection["EditorCC"] + "',JBM_IMMail='" + formCollection["IssueManagerEmail"] + "',JBM_RevProof='" + formCollection["ReviseseProofReq"] + "',PeRole='" + formCollection["PERole"] + "',PrName='" + formCollection["PRName"] + "',PrEmail='" + formCollection["PREmail"] + "',PrRole='" + formCollection["PRRole"] + "',EditorName='" + formCollection["EName"] + "',PrAdminName='" + formCollection["PAName"] + "',PrAdminEmail='" + formCollection["PAEmail"] + "',PrAdminRole='" + formCollection["ProofAdminRole"] + "',ObserverName='" + formCollection["OName"] + "'," +
"ObserverEmail='" + formCollection["OEmail"] + "',ObserverRole='" + formCollection["ObserverRole"] + "',ReviewOrder='" + formCollection["EPTParallel"] + "' ,OPS_wf='" + formCollection["OPS_wf"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalRate(FormCollection formCollection)
        {
            try
            {
                string strAutoIdFilter = Session["sCustAcc"].ToString();             

                string strQuery = @"update jbm_info set JBM_CeRate='" + formCollection["CERate"] + "',JBM_Rate='" + formCollection["JournalRate"] + "',IM_Rate='" + formCollection["IMRate"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalDeliverables(FormCollection formCollection)
        {
            try
            {
                string strAutoIdFilter = Session["sCustAcc"].ToString();
                formCollection["DView"] = formCollection["DView"].ToString() == "undefined" ? "" : formCollection["DView"];

                string strQuery = @"update jbm_info set Deliverables='" + formCollection["DView"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalWorkflow(FormCollection formCollection)
        {
            try
            {
                string strAutoIdFilter = Session["sCustAcc"].ToString();

                string strQuery = @"update jbm_info set Fp_wf='" + formCollection["Fp_wf"] + "',Rev_wf='" + formCollection["Rev_wf"] + "',Fin_wf='" + formCollection["Fin_wf"] + "',PapV_wf='" + formCollection["PapV_wf"] + "',PapA_wf='" + formCollection["PapA_wf"] + "',PapB_wf='" + formCollection["PapB_wf"] + "',Iss_wf='" + formCollection["Iss_wf"] + "',Onl_wf='" + formCollection["Onl_wf"] + "',SUP_wf='" + formCollection["SUP_wf"] + "',PRR_wf='" + formCollection["PRR_wf"] + "',Ver_wf='" + formCollection["Ver_wf"] + "' ,Pes_wf='" + formCollection["Pes_wf"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalIssue(FormCollection formCollection)
        {
            try
            {
                string strQuery = @"update jbm_info set FreshPageStart='" + formCollection["FreshPageStart"] + "' ,CompositePDF='" + formCollection["CompositePDF"] + "' ,ColorID='" + formCollection["Color"] + "',slug='" + formCollection["Slug"] + "',PageNo='" + formCollection["PageNo"] + "',CuttingMark='" + formCollection["CuttingMark"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalOthers(FormCollection formCollection)
        {
            try
            {
                formCollection["rSheridan"] = formCollection["rSheridan"].ToString() == "undefined" ? "" : formCollection["rSheridan"];
                formCollection["rSGMLOld"] = formCollection["rSGMLOld"].ToString() == "undefined" ? "" : formCollection["rSGMLOld"];

                string strQuery = @"update jbm_info set Sheridan='" + formCollection["rSheridan"] + "',SGMLOld_wf='" + formCollection["rSGMLOld"] + "',PageFormat='" + formCollection["Pageformat"] + "',Currency='" + formCollection["Currency"] + "',JNL_GSM='" + formCollection["GSM"] + "',JBM_SmallFormat='" + formCollection["LargeFormat"] + "',EarlyXML='" + formCollection["EarlyXML"] + "',onl_rel='" + formCollection["PrvParallel"] + "',JTF_S2008='" + formCollection["JTF_S2008"] + "',JTF_Informa='" + formCollection["Informa"] + "',jSAM='" + formCollection["SAM"] + "',SAMAutEd='" + formCollection["SAMAutEd"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalInfo(FormCollection formCollection)
        {
            try
            {

                formCollection["PERole"] = formCollection["PERole"].ToString() == "" ? "" : formCollection["PERole"].Replace(",", "|") + "|";
                formCollection["PRRole"] = formCollection["PRRole"].ToString() == "" ? "" : formCollection["PRRole"].Replace(",", "|") + "|";
                formCollection["EditorRole"] = formCollection["EditorRole"].ToString() == "" ? "" : formCollection["EditorRole"].Replace(",", "|") + "|";
                formCollection["SeniorEditorRole"] = formCollection["SeniorEditorRole"].ToString() == "" ? "" : formCollection["SeniorEditorRole"].Replace(",", "|") + "|";
                formCollection["ProofAdminRole"] = formCollection["ProofAdminRole"].ToString() == "" ? "" : formCollection["ProofAdminRole"].Replace(",", "|") + "|";
                formCollection["ObserverRole"] = formCollection["ObserverRole"].ToString() == "" ? "" : formCollection["ObserverRole"].Replace(",", "|") + "|";
                string strAutoIdFilter = Session["sCustAcc"].ToString();
                formCollection["CWPURL"] = formCollection["CWPURL"].ToString() == "undefined" ? "" : formCollection["CWPURL"];
                formCollection["DView"] = formCollection["DView"].ToString() == "undefined" ? "" : formCollection["DView"];
                formCollection["rSheridan"] = formCollection["rSheridan"].ToString() == "undefined" ? "" : formCollection["rSheridan"];
                formCollection["rSGMLOld"] = formCollection["rSGMLOld"].ToString() == "undefined" ? "" : formCollection["rSGMLOld"];
                formCollection["DelyID"] = formCollection["DelyID"].ToString() == "[Select]" ? "" : formCollection["DelyID"];

                string strQuery = @"update jbm_info set JBM_Intrnl='" + formCollection["InternalID"] + "',JBM_ID='" + formCollection["InternalID"] + "',JBM_Name='" + formCollection["JournalName"] + "',Title='" + formCollection["Title"] + "',JBM_IssueMan='"
+ formCollection["IssueManager"] + "',FTPID='" + formCollection["FTPID"] + "',JBM_Disabled='" + formCollection["Disabled"] + "',JBM_mFrom='" + formCollection["FromEmail"] + "',JBM_mEditor='" + formCollection["EEmail"] + "',JBM_mReplyto='" + formCollection["ReplyEmail"] + "',JBM_mRemainder='" + formCollection["Authorchaserdays"] + "'," +
"JBM_mBCc='" + formCollection["BCC"] + "',JBM_mExpire='" + formCollection["Proofexpire"] + "',JBM_Im='" + formCollection["IMJournal"] + "',JBM_Level='" + formCollection["LevelID"] + "',JBM_CeRate='" + formCollection["CERate"] + "',JBM_Rate='" + formCollection["JournalRate"] + "',JBM_Sproof='" + formCollection["EProofReq"] + "'" +
",JBM_ProofType='" + formCollection["ProofType"] + "',Deliverables='" + formCollection["DView"] + "',Sheridan='" + formCollection["rSheridan"] + "',SGMLOld_wf='" + formCollection["rSGMLOld"] + "',JBM_Platform='" + formCollection["PlatformID"] + "',JBM_Trimsize='" + formCollection["TrimSize"] + "',CustID='" + formCollection["CustID"] + "',JBM_PeName='" + formCollection["PEName"] + "'" +
",JBM_mPe='" + formCollection["PEEmail"] + "',JBM_CustPe='" + formCollection["CustomerPEEmail"] + "',JBM_mCc='" + formCollection["CC"] + "',JBM_ProdMan='" + formCollection["PMName"] + "',JBM_ProdManEmail='" + formCollection["PMEmail"] + "',JBM_ProdManTele='" + formCollection["PMTele"] + "',JBM_PeRemainder='" + formCollection["PERoleChaserDay"]
+ "',JBM_PrRemainder='" + formCollection["PRRoleChaserDay"] + "',JBM_EdtRemainder='" + formCollection["ERoleChaserDay"] + "',JBM_eProofAttachment='" + formCollection["MailAttachment"] + "',IM_INPUT_JID='" + formCollection["IMInputJID"] + "',HWCODE='" + formCollection["HWCode"] + "',DelyID='" + formCollection["DelyID"]
+ "',POD_required='" + formCollection["PrintOnDelivery"] + "',PrePrintJrnl='" + formCollection["PreprintJournal"] + "',Composition='" + formCollection["Composition"] + "',RevEpt_Wf='" + formCollection["RevisesEPTWF"] + "',JBM_IgnoreEforms='" + formCollection["IgnoreEForms"] + "',JBM_IgnoreForms='" +
formCollection["IgnoreForms"] + "',JBM_MailFormatHTML='" + formCollection["MailFormatHtml"] + "',JBM_IgnoreEproofCmtEnable='" + formCollection["IgnoreEproofCmtEnable"] + "',JBM_WaterMarkLogo='" + formCollection["WaterMarkLogo"] + "',JBM_WaterMarkLogoSettings='" + formCollection["WMLogoSettings"] + "',JBM_WaterMarkNA='" +
formCollection["WaterMarkNA"] + "',JBM_FormsPdfRuntime='" + formCollection["FormPDFRuntime"] + "',JBM_GenEProofTpl='" + formCollection["GenEProof"] + "',JBM_IgnoreArticleInProofAttachment='" + formCollection["IgnoreArticleInProofAttch"] + "',CWPURL='" + formCollection["CWPURL"] + "',OUP_Code='" +
formCollection["OUPCode"] + "',website='" + formCollection["JournalWebsite"] + "',DOI_prefix='" + formCollection["DOIPrefix"] + "',IM_Rate='" + formCollection["IMRate"] + "',JBM_mRemCc='" + formCollection["ReminderCC"] + "',Observer_MailMode='" + formCollection["CWPBccMail"] + "',JBM_mSrEditor='" + formCollection["SEEmail"] + "',JBM_SrEdtRemainder='" +
formCollection["SERoleChaserDay"] + "',SrEditorName='" + formCollection["SEName"] + "',EditorRole='" + formCollection["EditorRole"] + "',SrEditorRole='" + formCollection["SeniorEditorRole"] + "',JBM_Editor2Name='" + formCollection["Editor2Name"] + "',JBM_mEditor2_To='" + formCollection["Editor2Email"] + "'" +
",JBM_mEditor2_Cc='" + formCollection["Editor2CC"] + "',CcEditor='" + formCollection["EditorCC"] + "',OnlOnly='" + formCollection["OnlineOnly"] + "',JBM_IMMail='" + formCollection["IssueManagerEmail"] + "',JBM_RevProof='" + formCollection["ReviseseProofReq"] + "',JBM_TeamID='" + formCollection["Team"] + "',JBM_CeRequired='" + formCollection["CopyEditing"] + "',JBM_CeRequiredOffShore='" + formCollection["CopyEditingoffshore"] + "',JNL_ISSNPrint='" + formCollection["PrintISSN"] + "'" +
",JNL_ISSNonline='" + formCollection["OnlineISSN"] + "',JBM_SngDbl='" + formCollection["Column"] + "',PE_Approval='" + formCollection["PEApproval"] + "',PAP_A='" + formCollection["ModelAPAP"] + "',PAP_B='" + formCollection["ModelBPAP"] + "',PE_ApprovalWait='" + formCollection["WaitingPE"] + "',JBM_PrinterName='" + formCollection["PrinterName"] + "',PageExtent='" + formCollection["PageExtent"] + "'," +
"RevisesChaserDays='" + formCollection["RevisesAuthorChasing"] + "',FrequentIssuePerYear='" + formCollection["Frequency"] + "',FtpLoc='" + formCollection["USUKJournals"] + "',Dictionary='" + formCollection["Dictionary"] + "',Copyrightowner='" + formCollection["CopyRightOwner"] + "',OPS_Rev_wf='" + formCollection["RevisesSPWorkflow"] + "'," +
"PeRole='" + formCollection["PERole"] + "',PrName='" + formCollection["PRName"] + "',PrEmail='" + formCollection["PREmail"] + "',PrRole='" + formCollection["PRRole"] + "',EditorName='" + formCollection["EName"] + "',PrAdminName='" + formCollection["PAName"] + "',PrAdminEmail='" + formCollection["PAEmail"] + "',PrAdminRole='" + formCollection["ProofAdminRole"] + "',ObserverName='" + formCollection["OName"] + "'," +
"ObserverEmail='" + formCollection["OEmail"] + "',ObserverRole='" + formCollection["ObserverRole"] + "',ReviewOrder='" + formCollection["EPTParallel"] + "' ,FreshPageStart='" + formCollection["FreshPageStart"] + "' ,CompositePDF='" + formCollection["CompositePDF"] + "' ,ColorID='" + formCollection["Color"] + "',slug='" + formCollection["Slug"] + "',PageNo='" + formCollection["PageNo"] + "',CuttingMark='" + formCollection["CuttingMark"] + "'," +
"PageFormat='" + formCollection["Pageformat"] + "',Currency='" + formCollection["Currency"] + "',JNL_GSM='" + formCollection["GSM"] + "',JBM_SmallFormat='" + formCollection["LargeFormat"] + "',EarlyXML='" + formCollection["EarlyXML"] + "',onl_rel='" + formCollection["PrvParallel"] + "',JTF_S2008='" + formCollection["JTF_S2008"] + "',JTF_Informa='" + formCollection["Informa"] + "',jSAM='" + formCollection["SAM"] + "',SAMAutEd='" + formCollection["SAMAutEd"] + "'," +
"Fp_wf='" + formCollection["Fp_wf"] + "',Rev_wf='" + formCollection["Rev_wf"] + "',Fin_wf='" + formCollection["Fin_wf"] + "',PapV_wf='" + formCollection["PapV_wf"] + "',PapA_wf='" + formCollection["PapA_wf"] + "',PapB_wf='" + formCollection["PapB_wf"] + "',Iss_wf='" + formCollection["Iss_wf"] + "',Onl_wf='" + formCollection["Onl_wf"] + "',SUP_wf='" + formCollection["SUP_wf"] + "',PRR_wf='" + formCollection["PRR_wf"] + "',Ver_wf='" + formCollection["Ver_wf"] + "' ,Pes_wf='" + formCollection["Pes_wf"] + "',OPS_wf='" + formCollection["OPS_wf"] + "' where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        [HttpPost]
        public ActionResult UpdateJournalInf(FormCollection formCollection)
        {
            try
            {
                string strAutoIdFilter = Session["sCustAcc"].ToString();
                formCollection["DelyID"] = formCollection["DelyID"].ToString() == "[Select]" ? "" : formCollection["DelyID"];

                string strQuery = @"update jbm_info set JBM_Intrnl='" + formCollection["InternalID"] + "',JBM_ID='" + formCollection["InternalID"] + "',JBM_Name='" + formCollection["JournalName"] + "',Title='" + formCollection["Title"] + "',JBM_IssueMan='"
+ formCollection["IssueManager"] + "',FTPID='" + formCollection["FTPID"] + "',DelyID='" + formCollection["DelyID"]+"',JBM_Disabled ='" + formCollection["Disabled"] + "'," +
"JBM_Im='" + formCollection["IMJournal"] + "',JBM_Level='" + formCollection["LevelID"] + "',DOI_prefix='" + formCollection["DOIPrefix"] + "'," +
",JBM_Platform='" + formCollection["PlatformID"] + "',JBM_Trimsize='" + formCollection["TrimSize"] + "',CustID='" + formCollection["CustID"] + "'," +
"IM_INPUT_JID='" + formCollection["IMInputJID"] + "',HWCODE='" + formCollection["HWCode"] + "',website='" + formCollection["JournalWebsite"] + "'," +
"POD_required='" + formCollection["PrintOnDelivery"] + "',PrePrintJrnl='" + formCollection["PreprintJournal"] + "',Composition='" + formCollection["Composition"] + "',RevEpt_Wf='" + formCollection["RevisesEPTWF"] + "'," +
"OUP_Code='" +formCollection["OUPCode"] + "',OnlOnly='" + formCollection["OnlineOnly"] + "',JBM_TeamID='" + formCollection["Team"] + "',JBM_CeRequired='" + formCollection["CopyEditing"] + "',JBM_CeRequiredOffShore='" + formCollection["CopyEditingoffshore"] + "',JNL_ISSNPrint='" + formCollection["PrintISSN"] + "'" +
",JNL_ISSNonline='" + formCollection["OnlineISSN"] + "',JBM_SngDbl='" + formCollection["Column"] + "',PE_Approval='" + formCollection["PEApproval"] + "',PAP_A='" + formCollection["ModelAPAP"] + "',PAP_B='" + formCollection["ModelBPAP"] + "',PE_ApprovalWait='" + formCollection["WaitingPE"] + "',JBM_PrinterName='" + formCollection["PrinterName"] + "',PageExtent='" + formCollection["PageExtent"] + "'," +
"RevisesChaserDays='" + formCollection["RevisesAuthorChasing"] + "',FrequentIssuePerYear='" + formCollection["Frequency"] + "',FtpLoc='" + formCollection["USUKJournals"] + "',Dictionary='" + formCollection["Dictionary"] + "',Copyrightowner='" + formCollection["CopyRightOwner"] + "',OPS_Rev_wf='" + formCollection["RevisesSPWorkflow"] + "'," +
" where JBM_AutoID='" + Session["sJBM_AutoID"].ToString() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                return Json(new { data = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public string CreateBtn(string uniqueID)
        {
            string formControl = string.Empty;
            try
            {
                if (uniqueID != "")
                {
                    //formControl = "<button type='button' id='Edit" + uniqueID + "' onClick=\"compoEdit('" + uniqueID + "')\"  class='btn btn-light' name='edit' value='Edit'><span class='fas fa-edit fa-1x text-green'></span></button><button type='button' id='Delete" + uniqueID + "' onClick=\"compoDelete('" + uniqueID + "')\"  class='btn btn-light' name='delete' value='Delete'><span class='fas fa-trash fa-1x text-red'></span></button>";
                    formControl = " <a id='btnEdit" + uniqueID + "' href='javascript: void(0);' onClick=\"gotoEditDetails('" + uniqueID + "');\"><i class='far fa-edit' style='color:#28a745;font-size:14px'></i></a><a id = 'btndelete" + uniqueID + "' href='javascript:void(0);' onclick=\"gotoEditDetails('" + uniqueID + "');\"><i class='far fa-trash-alt' style='color:#28a745;font-size:14px'></i></a>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }



        [SessionExpire]
        public ActionResult AddNewJAT(string sArtTypeMTFoot, string sSubArtTypeFoot)
        {
            try
            {
                if (sArtTypeMTFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";
                    DataTable dt = DBProc.GetResultasDataTbl("Select ArtTypeID from JBM_JrnlArticleTypes Where JBM_AutoId='" + strJAutoId + "' and ArtTypeID='" + sArtTypeMTFoot + "'", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                    }
                    string strQuery = "INSERT INTO JBM_JrnlArticleTypes ([JBM_AutoID], [ArtTypeID],[ArticleType]) values('" + strJAutoId + "','" + sArtTypeMTFoot + "','" + sSubArtTypeFoot + "')";
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
        public ActionResult UpdateJAT(string sArtTypeMTFoot, string sSubArtTypeFoot, string sPAP, string sLIC)
        {
            try
            {
                if (sArtTypeMTFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "Update JBM_JrnlArticleTypes set Pap_Required='" + sPAP + "',License_Required='" + sLIC + "' ,ArticleType='" + sSubArtTypeFoot + "' where jbm_AutoID='" + strJAutoId + "' and arttypeid='" + sArtTypeMTFoot + "'";
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
        public ActionResult DeleteJAT(string sArtTypeMTFoot)
        {
            try
            {
                if (sArtTypeMTFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "DELETE FROM JBM_JrnlArticleTypes where jbm_AutoID='" + strJAutoId + "' and arttypeid='" + sArtTypeMTFoot + "'";
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
        public ActionResult AddNewTAT(string sArtTypeTATFoot, string sPriority, string sTAT, string sLogin, string sVendor, string sRE, string sCE, string sComp)
        {
            try
            {
                if (sArtTypeTATFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";
                    DataTable dt = DBProc.GetResultasDataTbl("Select JBM_AutoID from JBM_TAT_Master Where JBM_AutoId='" + strJAutoId + "' and stage='" + sArtTypeTATFoot + "' and Priority='" + sPriority + "'", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                    }
                    string strQuery = "INSERT INTO JBM_TAT_Master ([JBM_AutoID], [Stage], [Priority], [TAT], [LoginTAT], [VendorTAT], [ReTAT],[CeTAT], [CompTAT]) VALUES ('" + strJAutoId + "', '" + sArtTypeTATFoot + "', '" + sPriority + "', '" + sTAT + "', '" + sLogin + "', '" + sVendor + "', '" + sRE + "', '" + sCE + "', '" + sComp + "')";
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
        public ActionResult UpdateTAT(string sArtTypeTATFoot, string sPriority, string sTAT, string sLogin, string sVendor, string sRE, string sCE, string sComp, string strAutoID)
        {
            try
            {
                if (sArtTypeTATFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "UPDATE JBM_TAT_Master  SET [Stage] = '" + sArtTypeTATFoot + "', [Priority] = '" + sPriority + "', [TAT] = '" + sTAT + "', [LoginTAT] = '" + sLogin + "', [VendorTAT] = '" + sVendor + "', [ReTAT] = '" + sRE + "', [CeTAT] = '" + sCE + "', [CompTat] = '" + sComp + "'  WHERE [AutoId] = '" + strAutoID + "'";
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
        public ActionResult DeleteTAT(string strAutoID)
        {
            try
            {
                if (strAutoID.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "DELETE FROM JBM_TAT_Master WHERE AutoID='" + strAutoID + "'";
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
        public ActionResult SaveTAT(string CheckedStageID)
        {
            try
            {
                if (CheckedStageID.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";
                    List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedStageID);
                    if (chkIds.Count > 0)
                    {
                        for (int i = 0; i < chkIds.Count; i++)
                        {
                            string strPriority = chkIds[i].Split('|')[0];
                            string strTAT = chkIds[i].Split('|')[1];
                            string strLogin = chkIds[i].Split('|')[2];
                            string strVendor = chkIds[i].Split('|')[3];
                            string strRE = chkIds[i].Split('|')[4];
                            string strCE = chkIds[i].Split('|')[5];
                            string strComp = chkIds[i].Split('|')[6];
                            string IncSat = chkIds[i].Split('|')[7];
                            string IncSun = chkIds[i].Split('|')[8];
                            string strTATR = chkIds[i].Split('|')[9];
                            string strLoginR = chkIds[i].Split('|')[10];
                            string strVendorR = chkIds[i].Split('|')[11];
                            string strRER = chkIds[i].Split('|')[12];
                            string strCER = chkIds[i].Split('|')[13];
                            string strCompR = chkIds[i].Split('|')[14];
                            string IncSatR = chkIds[i].Split('|')[15];
                            string IncSunR = chkIds[i].Split('|')[16];
                            string strQuery = "";
                            string strResult2 = "";
                            DataTable dt = DBProc.GetResultasDataTbl("select jbm_Autoid from JBM_TAT_Master where JBM_AutoID='" + strJAutoId + "' and stage='FP' and Priority='" + strPriority + "'", Session["sConnSiteDB"].ToString());
                            if (dt.Rows.Count == 0)
                            {
                                strQuery = "Insert into JBM_TAT_Master (JBM_AutoID,Stage,Priority,TAT,LoginTAT,VendorTAT,ReTAT,CeTAT,CompTat,IncSaturday,IncSunday) Values('" + strJAutoId + "','FP','" + strPriority + "','" + strTAT + "','" + strLogin + "','" + strVendor + "','" + strRE + "','" + strCE + "','" + strComp + "','" + IncSat + "','" + IncSun + "')";
                            }
                            else
                            {
                                strQuery = "Update JBM_TAT_Master set JBM_AutoID='" + strJAutoId + "',Stage='FP',Priority='" + strPriority + "',TAT='" + strTAT + "',LoginTAT='" + strLogin + "',VendorTAT='" + strVendor + "',ReTAT='" + strRE + "',CeTAT='" + strCE + "',CompTat='" + strComp + "',IncSaturday='" + IncSat + "',IncSunday='" + IncSun + "' where JBM_AutoID='" + strJAutoId + "' and stage='FP' and Priority='" + strPriority + "'";
                            }
                            strResult2 = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                            dt = DBProc.GetResultasDataTbl("select jbm_Autoid from JBM_TAT_Master where JBM_AutoID='" + strJAutoId + "' and stage='Rev' and Priority='" + strPriority + "'", Session["sConnSiteDB"].ToString());
                            if (dt.Rows.Count == 0)
                            {
                                strQuery = "Insert into JBM_TAT_Master (JBM_AutoID,Stage,Priority,TAT,LoginTAT,VendorTAT,ReTAT,CeTAT,CompTat,IncSaturday,IncSunday) Values('" + strJAutoId + "','Rev','" + strPriority + "','" + strTATR + "','" + strLoginR + "','" + strVendorR + "','" + strRER + "','" + strCER + "','" + strCompR + "','" + IncSatR + "','" + IncSunR + "')";
                            }
                            else
                            {
                                strQuery = "Update JBM_TAT_Master set JBM_AutoID='" + strJAutoId + "',Stage='Rev',Priority='" + strPriority + "',TAT='" + strTATR + "',LoginTAT='" + strLoginR + "',VendorTAT='" + strVendorR + "',ReTAT='" + strRER + "',CeTAT='" + strCER + "',CompTat='" + strCompR + "',IncSaturday='" + IncSatR + "',IncSunday='" + IncSunR + "' where JBM_AutoID='" + strJAutoId + "' and stage='Rev' and Priority='" + strPriority + "'";
                            }
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
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }

        [SessionExpire]
        public ActionResult AddNewTYPE(string sArtTypeFoot, string sWorkflow, string sCC, string sBCC, string sEditorName, string sEditor, string sEditorCC, string sEditor2Name, string sEditor2, string sEditor2CC, string sPEName, string sPE, string sPECC, string sPRName, string sPR, string sPRCC)
        {
            try
            {
                if (sArtTypeFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";
                    DataTable dt = DBProc.GetResultasDataTbl("Select ArtTypeID from " + Init_Tables.gTblJBMArticleCategory + " Where JBM_AutoId='" + strJAutoId + "' and ArtTypeId='" + sArtTypeFoot + "'", Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count > 0)
                    {
                        return Json(new { dataComp = "Exists" }, JsonRequestBehavior.AllowGet);
                    }
                    string strQuery = "INSERT INTO " + Init_Tables.gTblJBMArticleCategory + " ([JBM_AutoID], [OPS_wf],[CustType], [ArtTypeID],[mCc], [mBCc],[mEditorName],[mEditor2_Name],[mPeName],[mPrName],[mEditor],[mEditor2_To],[mPe],[mPr],[mEditor_cc],[mEditor2_cc],[mPe_cc],[mPr_cc]) VALUES ('" + strJAutoId + "', '" + sWorkflow + "', '" + Session["sCustAcc"].ToString() + "', '" + sArtTypeFoot + "', '" + sCC + "','" + sBCC + "' ,'" + sEditorName + "','" + sEditor2Name + "','" + sPEName + "','" + sPRName + "','" + sEditor + "','" + sEditor2 + "','" + sPE + "','" + sPR + "','" + sEditorCC + "','" + sEditor2CC + "','" + sPECC + "','" + sPRCC + "')";
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
        public ActionResult UpdateTYPE(string sArtTypeFoot, string sWorkflow, string sCC, string sBCC, string sEditorName, string sEditor, string sEditorCC, string sEditor2Name, string sEditor2, string sEditor2CC, string sPEName, string sPE, string sPECC, string sPRName, string sPR, string sPRCC, string strAutoID)
        {
            try
            {
                if (sArtTypeFoot.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "UPDATE p SET [ArtTypeID] = '" + sArtTypeFoot + "', p.mCc='" + sCC + "',p.mBCc='" + sBCC + "',p.mEditorName='" + sEditorName + "',p.mEditor2_Name='" + sEditor2Name + "',p.mPeName='" + sPEName + "',p.mPrName='" + sPRName + "',p.mEditor='" + sEditor + "',p.mEditor2_To='" + sEditor2 + "',p.mPe='" + sPE + "',p.mPr='" + sPR + "',p.mEditor_cc='" + sEditorCC + "',p.mEditor2_cc='" + sEditor2CC + "',p.mPe_cc='" + sPECC + "',p.mPr_cc='" + sPRCC + "',p.OPS_WF='" + sWorkflow + "'  from JBM_ArticleCategory p  WHERE [AutoId] = '" + strAutoID + "'";
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
        public ActionResult DeleteTYPE(string strAutoID)
        {
            try
            {
                if (strAutoID.ToString() != "")
                {
                    string strJAutoId = "";
                    if (Session["sJBM_AutoID"] != null)
                        strJAutoId = Session["sJBM_AutoID"].ToString().Trim();
                    else
                        strJAutoId = "";

                    string strQuery = "DELETE FROM " + Init_Tables.gTblJBMArticleCategory + " WHERE AutoID='" + strAutoID + "'";
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