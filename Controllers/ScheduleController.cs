using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using SmartTrack.Helper;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using RL = ReferenceLibrary;
using System.Xml;
using System.IO.Compression;
using System.Data.OleDb;

namespace SmartTrack.Controllers
{
    [SessionExpire]
    public class ScheduleController : Controller
    {
        Dictionary<string, string> dicCollection = new Dictionary<string, string>();
        clsCollection clsCollec = new clsCollection();
        DataProc DBProc = new DataProc(); // Data store/retrive DB

        // GET: Schedule
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult SageSchedule() //string CustAcc, string EmpID, string SiteID, string URL
        {
            //if (URL == null)
            //{
            //    Session["returnURL"] = "../Schedule/SageSchedule";
            //}
            //else { Session["returnURL"] = URL; }

            //DataTable dtnew = new DataTable();
            //if (CustAcc != null)
            //{
            //    Session["EmpIdLogin"] = EmpID;
            //    Session["sCustAcc"] = CustAcc;
            //    Session["sSiteID"] = SiteID;
            //    clsCollec.getSiteDBConnection(SiteID, CustAcc);

            //    if (Session["sConnSiteDB"] == null || Session["sConnSiteDB"].ToString() == "")
            //    {
            //        Session["sConnSiteDB"] = GlobalVariables.strConnSite;
            //    }
            //    DataSet ds = new DataSet();
            //    ds = DBProc.GetResultasDataSet("Select EmpLogin,EmpAutoId,EmpName,EmpMailId,DeptCode,RoleId, JwAccessItm,TeamID, DeptAccess,CustAccess,TeamID,(Select b.DeptName from JBM_DepartmentMaster b  where e.DeptCode = b.DeptCode) as DeptName from JBM_EmployeeMaster e Where EmpLogin='" + EmpID + "'", Session["sConnSiteDB"].ToString());
            //    if (ds.Tables[0].Rows.Count > 0)
            //    {
            //        Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
            //        Session["EmpLogin"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
            //        Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();
            //        Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
            //        Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
            //        Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();
            //        Session["gJwAccItm"] = ds.Tables[0].Rows[0]["JwAccessItm"].ToString();
            //        Session["sCustGroup"] = "";
            //        Session["sCustTeamID"] = "";
            //        Session["gTeamID"] = ds.Tables[0].Rows[0]["TeamID"].ToString();
            //    }

            //}


            ViewBag.JournalList = getJournalList();
            ViewBag.ResourceList = getResourceList();
            return View();
        }
        public ActionResult getArticleLists(string vJBMID)
        {
            try
            {

                string strQueryFinal = "Select AutoArtID,ChapterID from " + Session["sCustAcc"].ToString() + "_ArticleInfo where JBM_AutoID='" + vJBMID + "'";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by ChapterID asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString()
                     };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult getArticleSchedule(string vJBMID, string vAutoArtID)
        {
            try
            {
                string strQueryFinal = "Select SI.AutoID,SI.JBM_AutoID,SI.AutoArtID,SI.TaskName,SI.Task_TAT,SI.Plan_StartDate,SI.Plan_EndDate,SI.Status,SI.ResourceName,(Select EmpName from JBM_EmployeeMaster Where EmpAutoID=SI.PlannedBy) as [PlannedBy] from JBM_ScheduleInfo SI WHERE SI.AutoArtID='" + vAutoArtID + "'";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by AutoID asc ", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {
                                     CreateBtn(a[0].ToString(), a[3].ToString(), a[2].ToString(), "TaskName"),
                                     CreateBtn(a[0].ToString(), a[4].ToString(), a[2].ToString(), "TaskDays"),
                                     //CreateBtn(a[0].ToString(), a[5].ToString(), a[2].ToString(), "StartDate"),
                                     //CreateBtn(a[0].ToString(), a[6].ToString(), a[2].ToString(), "EndDate"),
                                     CreateBtn(a[0].ToString(), a[8].ToString(), a[2].ToString(), "ResourceList"),
                                     a[7].ToString(),
                                     a[9].ToString()
                 };

                return Json(new { dataSch = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public string CreateBtn(string strAutoID, string strField, string strAutoArtID, string type)
        {
            string formControl = string.Empty;
            try
            {
                if (type == "StartDate" || type == "EndDate")
                {
                    formControl = "<input type='text' id='" + strAutoID + "_" + strAutoArtID + type + "' class='form-control datepicker date ht' value='" + strField + "'>";
                }
                else if (type == "TaskDays")
                {
                    formControl = "<input type='text' id='" + strAutoID + "_" + strAutoArtID + type + "' onclick=\"funcInputValidate()\" class='form-control ht' value='" + strField + "'>";
                }
                else if (type == "TaskName")
                {
                    formControl = "<input type='checkbox' class='caseChk' id='" + strAutoID + "'  onclick=\"funcCheckItem('" + strAutoID + "')\"  class='form-control text-center' name='" + strAutoID + "' value='KGLs'>&nbsp;&nbsp;<label for='" + strAutoID + "'><span id='" + strAutoID + "_" + strAutoArtID + type + "'>" + strField + "</span></label>"; // onclick=\"funcCheckItem('" + strAutoID + "')\" 
                }
                else if (type == "ResourceList")
                {
                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select ResourceID,ResourceName from JBM_ResourceMaster", Session["sConnSiteDB"].ToString());

                    formControl = "<select id='" + strAutoID + "_" + strAutoArtID + type + "' class='form-control ht' style='text-align:left;width:250px;'><option value=''>[Select]</option>[[ReplaceList]]</select>";

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string lstColl = string.Empty;
                        for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                        {
                            string ResourceID = ds.Tables[0].Rows[intCount]["ResourceID"].ToString();
                            string ResourceName = ds.Tables[0].Rows[intCount]["ResourceName"].ToString();

                            if (strField.Trim() == ResourceID.Trim())
                            {
                                lstColl += "<option value='" + ResourceID.Trim() + "' selected='selected'>" + ResourceName.Trim() + "</option>";
                            }
                            else
                            {
                                lstColl += "<option value='" + ResourceID.Trim() + "'>" + ResourceName.Trim() + "</option>";
                            }


                        }
                        formControl = formControl.Replace("[[ReplaceList]]", lstColl);
                    }
                    else
                    {
                        formControl = formControl.Replace("[[ReplaceList]]", "");
                    }

                    // formControl = "<select id='" + type + strAutoID + "' class='form-control ht' style='text-align:left'><option value=''>[Select]</option><option value='G2065'>101985</option><option selected='selected' value='G0921'>101987</option><option value='G0922'>367359</option></select>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        private List<SelectListItem> getJournalList(string isActive = "1")
        {
            List<SelectListItem> list = new List<SelectListItem>();
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet("Select JBM_AutoID,JBM_ID from JBM_Info where CustId in ('C002','C022') and JBM_Disabled=0", Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                {
                    list.Add(new SelectListItem
                    {
                        Text = ds.Tables[0].Rows[intCount]["JBM_ID"].ToString(),
                        Value = ds.Tables[0].Rows[intCount]["JBM_AutoID"].ToString()
                    });
                }

            }
            return list;

        }
        private List<SelectListItem> getResourceList(string isActive = "1")
        {
            List<SelectListItem> list = new List<SelectListItem>();
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet("Select ResourceID,ResourceName from JBM_ResourceMaster", Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                {
                    list.Add(new SelectListItem
                    {
                        Text = ds.Tables[0].Rows[intCount]["ResourceID"].ToString().Trim(),
                        Value = ds.Tables[0].Rows[intCount]["ResourceName"].ToString().Trim()
                    });
                }

            }
            return list;

        }
        public ActionResult UpdateArticleMetadataEvent(string vJBMID, string vAutoArtID, string vArticleID, string vEventType)
        {
            // Dictionary<string, string> dicCollection = new Dictionary<string, string>();
            try
            {
                string strFunderInfoCollection = string.Empty;
                string strAuthorInfoCollection = string.Empty;

                RL.ArtDet A = new RL.ArtDet();

                A.AD.JAutoID = vJBMID;
                A.AD.AutoArtID = vAutoArtID;
                A.AD.ArticleID = vArticleID;

                string strSubTitle = string.Empty;
                string strOtherArticleType = string.Empty;
                string strSpecialIssueOrCollectionsName = string.Empty;
                string blnIsExcludeFromOnlineFirst = "false";
                int intWordCount = 0;
                int intFigCount = 0;
                int intTableCount = 0;
                string blnIsLatex = "false";
                string blnIsColor = "false";
                string blnIsColorPayment = "false";
                string strManuscriptID = string.Empty; string strManuscriptType = string.Empty;
                string strSubmittedDate = string.Empty; string strFundingInfo = string.Empty;
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select (Select ArtTypeDesc from JBM_ArticleTypes WHERE ArtTypeID=AI.ArtTypeID) as [ArticleType],AI.etpArticleType  as [OtherArticleType], P.ArticleTitle as [Title],AI.Short_Title as [SubTitle],AI.Iss as [Issue],P.IssueSection  as [IssueSection],P.SplIssue as [SpecialIssueOrCollectionsName],P.IsExcludeFromOnlineFirst as [IsExcludeFromOnlineFirst],AI.WordCount as [WordCount],AI.NumofFigures as [FigureCount],AI.NumofTables as [TableCount],AI.NumofMSP as [ActualPages],AI.PlatformID as [IsLatex],P.Color as [IsContainsColour],P.IsColourFiguresRequirePayment as [IsColourFiguresRequirePayment],AI.SuppInfo as [IsSupplementalMaterial],AI.ManuscriptID as [ManuscriptNumber],AI.ManusType as [ManuscriptType],CONVERT(varchar,P.JNAcceptedDate,23) as [AcceptanceDate],CONVERT(varchar,P.OUPSubmitted,23) as [SubmittedDate],CONVERT(varchar,P.JNRevisedDate,23) as [RevisedSubmissionDate],P.FundingInfo as [OpenFunders] from SG_ArticleInfo AI INNER JOIN SG_ProdInfo P ON AI.AutoArtID=P.AutoArtID WHERE AI.JBM_AutoID='" + vJBMID.Trim() + "' and AI.ChapterID='" + vArticleID.Trim() + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    A.AD.ArticleType = ds.Tables[0].Rows[0]["ArticleType"].ToString();
                    strOtherArticleType = ds.Tables[0].Rows[0]["OtherArticleType"].ToString();
                    if (strOtherArticleType.Length > 80)
                    { return Json(new { dataComp = "Other Article must be maximum of 80 characters." }, JsonRequestBehavior.AllowGet); }


                    A.AD.ArticleTitle = ds.Tables[0].Rows[0]["Title"].ToString();
                    if (A.AD.ArticleTitle == "" || A.AD.ArticleTitle == null)
                    { return Json(new { dataComp = "Title should not be empty." }, JsonRequestBehavior.AllowGet); }
                    else if (A.AD.ArticleTitle.Length > 850)
                    { return Json(new { dataComp = "Title must be maximum of 850 characters." }, JsonRequestBehavior.AllowGet); }

                    strSubTitle = ds.Tables[0].Rows[0]["SubTitle"].ToString();
                    if (strSubTitle.Length > 850)
                    { return Json(new { dataComp = "Sub Title must be maximum of 850 characters." }, JsonRequestBehavior.AllowGet); }

                    A.AD.iss = ds.Tables[0].Rows[0]["Issue"].ToString();
                    if (A.AD.iss.Length > 30)
                    { return Json(new { dataComp = "Issue must be maximum of 30 characters." }, JsonRequestBehavior.AllowGet); }

                    A.AD.IssType = ds.Tables[0].Rows[0]["IssueSection"].ToString();
                    if (A.AD.IssType.Length > 255)
                    { return Json(new { dataComp = "Issue Section must be maximum of 255 characters." }, JsonRequestBehavior.AllowGet); }

                    strSpecialIssueOrCollectionsName = ds.Tables[0].Rows[0]["SpecialIssueOrCollectionsName"].ToString();
                    if (strSpecialIssueOrCollectionsName.Length > 250)
                    { return Json(new { dataComp = "Special Issue or Collections Name must be maximum of 250 characters." }, JsonRequestBehavior.AllowGet); }

                    blnIsExcludeFromOnlineFirst = ds.Tables[0].Rows[0]["IsExcludeFromOnlineFirst"].ToString().ToLower().Trim();   // True or False
                    intWordCount = ds.Tables[0].Rows[0]["WordCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["WordCount"].ToString().Trim()) : 0;
                    intFigCount = ds.Tables[0].Rows[0]["FigureCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["FigureCount"].ToString().Trim()) : 0;
                    intTableCount = ds.Tables[0].Rows[0]["TableCount"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["TableCount"].ToString().Trim()) : 0;

                    A.AD.intActualPages = ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim() != "" ? Convert.ToInt32(ds.Tables[0].Rows[0]["ActualPages"].ToString().Trim()) : 0;
                    //if (A.AD.intActualPages == 0)
                    //{
                    //    { return Json(new { dataComp = "Actual pages should not be zero." }, JsonRequestBehavior.AllowGet); }
                    //}

                    if (ds.Tables[0].Rows[0]["IsLatex"].ToString().Trim() == "1")
                    { blnIsLatex = "true"; }
                    blnIsColor = ds.Tables[0].Rows[0]["IsContainsColour"].ToString().Trim() == "1" ? "true" : "false";
                    blnIsColorPayment = ds.Tables[0].Rows[0]["IsColourFiguresRequirePayment"].ToString().ToLower().Trim() == "1" ? "true" : "false";   // True or False
                    A.AD.SuppDataAvailable = ds.Tables[0].Rows[0]["IsSupplementalMaterial"].ToString().Trim() == "1" ? "true" : "false";
                    strManuscriptID = ds.Tables[0].Rows[0]["ManuscriptNumber"].ToString();
                    if (strManuscriptID.Length > 80)
                    { return Json(new { dataComp = "Manuscript Number must be maximum of 80 characters." }, JsonRequestBehavior.AllowGet); }

                    strManuscriptType = ds.Tables[0].Rows[0]["ManuscriptType"].ToString();
                    if (strManuscriptType.Length > 255)
                    { return Json(new { dataComp = "Manuscript Type must be maximum of 255 characters." }, JsonRequestBehavior.AllowGet); }

                    A.AD.JN_AcceptedDate = ds.Tables[0].Rows[0]["AcceptanceDate"].ToString();
                    strSubmittedDate = ds.Tables[0].Rows[0]["SubmittedDate"].ToString();
                    A.AD.JN_RevisedDate = ds.Tables[0].Rows[0]["RevisedSubmissionDate"].ToString();
                    strFundingInfo = ds.Tables[0].Rows[0]["OpenFunders"].ToString();

                }
                else {
                    { return Json(new { dataComp = "Article Id should not be empty." }, JsonRequestBehavior.AllowGet); }
                }

                string strValidResult = Proc_ContentValidation(vArticleID, A.AD.ArticleType, A.AD.ArticleTitle, blnIsExcludeFromOnlineFirst, intWordCount, intFigCount, intTableCount, A.AD.intActualPages);

                if (strValidResult != "")
                {
                    return Json(new { dataComp = strValidResult }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    /// Author information collection
                    Dictionary<string, string> primaryMaildic = new Dictionary<string, string>();
                    ds = new DataSet();
                    ds = DBProc.GetResultasDataSet("Select AutoArtID,Title as [Salutation],FirstName,LastName,Orcid_No as [ORCID],'' as [ArticleAffiliation],RinggoldId,Institution as [RinggoldInstitution], isCorrAuthor as [IsCorrespondingAuthor],EmailAddress as [PrimaryEmail], MailingAffiliation,Addressline1 as[Address1],Addressline2 as[Address2],City,State as [StateOrProvince],Country,Postalcode,IsGroupAuthor,GroupAuthorName,auth_seq as [SequenceNo] from SG_Gwpsuserinfo where AutoArtID='" + A.AD.AutoArtID + "' Order By [SequenceNo] ASC", Session["sConnSiteDB"].ToString());
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        strAuthorInfoCollection = string.Empty;
                        for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                        {
                            string blnCorrAuth = ds.Tables[0].Rows[0]["IsCorrespondingAuthor"].ToString().ToLower().Trim();//"false";

                            //if (ds.Tables[0].Rows[0]["IsCorrespondingAuthor"].ToString().ToLower().Trim() == "true")
                            //{ blnCorrAuth = "true"; }

                            string blnGroupAuth = ds.Tables[0].Rows[0]["IsGroupAuthor"].ToString().Trim() == "1" ? "true" : "false";  //"false";
                                                                                                                                      //if (ds.Tables[0].Rows[0]["IsGroupAuthor"].ToString().Trim() == "1")
                                                                                                                                      //{ blnGroupAuth = "true"; }

                            if (ds.Tables[0].Rows[intCount]["Salutation"].ToString().Length > 80)
                            { return Json(new { dataComp = "Salutation must be maximum of 80 characters." }, JsonRequestBehavior.AllowGet); }

                            if ((ds.Tables[0].Rows[intCount]["FirstName"].ToString() == "" || ds.Tables[0].Rows[intCount]["FirstName"].ToString() == null) && blnGroupAuth == "false")
                            { return Json(new { dataComp = "First name should not be empty." }, JsonRequestBehavior.AllowGet); }

                            if ((ds.Tables[0].Rows[intCount]["LastName"].ToString() == "" || ds.Tables[0].Rows[intCount]["LastName"].ToString() == null) && blnGroupAuth == "false")
                            { return Json(new { dataComp = "Last name should not be empty." }, JsonRequestBehavior.AllowGet); }

                            if (ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString() == "" || ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString() == null)
                            { return Json(new { dataComp = "Primary Email should not be empty." }, JsonRequestBehavior.AllowGet); }
                            if (!primaryMaildic.Keys.Contains(ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString()))
                            {
                                primaryMaildic.Add(ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString(), ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString());
                            }
                            else
                            { return Json(new { dataComp = "Primary email address should be unique for author list within the article." }, JsonRequestBehavior.AllowGet); }


                            if ((ds.Tables[0].Rows[intCount]["Address1"].ToString() == "" || ds.Tables[0].Rows[intCount]["Address1"].ToString() == null) && ds.Tables[0].Rows[intCount]["IsCorrespondingAuthor"].ToString() == "1" && (ds.Tables[0].Rows[intCount]["MailingAffiliation"].ToString() == "" || ds.Tables[0].Rows[intCount]["MailingAffiliation"].ToString() == null))
                            { return Json(new { dataComp = "Address1 should not be empty for corresponding author." }, JsonRequestBehavior.AllowGet); }
                            else if (ds.Tables[0].Rows[intCount]["Address1"].ToString().Length > 160)
                            { return Json(new { dataComp = "Address1 must be maximum of 160 characters." }, JsonRequestBehavior.AllowGet); }


                            if ((ds.Tables[0].Rows[intCount]["City"].ToString() == "" || ds.Tables[0].Rows[intCount]["City"].ToString() == null) && blnCorrAuth == "true")
                            { return Json(new { dataComp = "City should not be empty for corresponding author." }, JsonRequestBehavior.AllowGet); }

                            //Country validation
                            if ((ds.Tables[0].Rows[intCount]["Country"].ToString() == "" || ds.Tables[0].Rows[intCount]["Country"].ToString() == null) && blnCorrAuth == "true")
                            { return Json(new { dataComp = "A country should not be empty for name author." }, JsonRequestBehavior.AllowGet); }
                            else if (blnCorrAuth == "true" && ds.Tables[0].Rows[intCount]["Country"].ToString() != "")
                            {
                                DataSet dsCountry = new DataSet();
                                dsCountry = DBProc.GetResultasDataSet("Select CountryID from JBM_CountryMaster WHERE (CountryCode1='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "' or CountryCode2='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "')", Session["sConnSiteDB"].ToString());
                                if (dsCountry.Tables[0].Rows.Count == 0)
                                { return Json(new { dataComp = "A country provided for name author is not listed in SmartJ." }, JsonRequestBehavior.AllowGet); }
                            }

                            //State Province validation
                            DataSet dsState = new DataSet();
                            if ((ds.Tables[0].Rows[intCount]["StateOrProvince"].ToString() == "" || ds.Tables[0].Rows[intCount]["StateOrProvince"].ToString() == null) && ds.Tables[0].Rows[intCount]["Country"].ToString() != "" && blnCorrAuth == "true")
                            {
                                dsState = DBProc.GetResultasDataSet("Select CountryStateCode from JBM_StateMaster WHERE CountryID=(Select CountryID from JBM_CountryMaster WHERE (CountryCode1='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "' or CountryCode2='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "'))", Session["sConnSiteDB"].ToString());
                                if (dsState.Tables[0].Rows.Count > 0)
                                { return Json(new { dataComp = "A state or province should not be empty for name author." }, JsonRequestBehavior.AllowGet); }
                            }
                            else if (blnCorrAuth == "true" && ds.Tables[0].Rows[intCount]["StateOrProvince"].ToString() != "")
                            {
                                dsState = DBProc.GetResultasDataSet("Select CountryStateCode from JBM_StateMaster WHERE CountryID=(Select CountryID from JBM_CountryMaster WHERE (CountryCode1='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "' or CountryCode2='" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "')) and CountryStateCode='" + ds.Tables[0].Rows[intCount]["StateOrProvince"].ToString() + "'", Session["sConnSiteDB"].ToString());
                                if (dsState.Tables[0].Rows.Count == 0)
                                { return Json(new { dataComp = "A state provided for name author is not listed in SmartJ." }, JsonRequestBehavior.AllowGet); }
                            }


                            if ((ds.Tables[0].Rows[intCount]["Postalcode"].ToString() == "" || ds.Tables[0].Rows[intCount]["Postalcode"].ToString() == null) && blnCorrAuth == "true" && Regex.IsMatch(ds.Tables[0].Rows[intCount]["Country"].ToString(), "(CA|US)",RegexOptions.IgnoreCase))
                            {
                                return Json(new { dataComp = "A postal code should not be empty for corresponding author." }, JsonRequestBehavior.AllowGet);
                            }

                            if (ds.Tables[0].Rows[intCount]["SequenceNo"].ToString() == "" || ds.Tables[0].Rows[intCount]["SequenceNo"].ToString() == null)
                            { return Json(new { dataComp = "Sequence no should not be empty." }, JsonRequestBehavior.AllowGet); }
                            else if (ds.Tables[0].Rows[intCount]["SequenceNo"].ToString() == "0")
                            { return Json(new { dataComp = "Sequence no should be valid number." }, JsonRequestBehavior.AllowGet); }


                            if ((ds.Tables[0].Rows[intCount]["GroupAuthorName"].ToString() == "" || ds.Tables[0].Rows[intCount]["GroupAuthorName"].ToString() == null) && blnGroupAuth == "true")
                            { return Json(new { dataComp = "Group name should not be empty." }, JsonRequestBehavior.AllowGet); }

                            strAuthorInfoCollection += "{\"Salutation\": \"" + ds.Tables[0].Rows[intCount]["Salutation"].ToString() + "\",\"FirstName\": \"" + ds.Tables[0].Rows[intCount]["FirstName"].ToString() + "\", \"LastName\": \"" + ds.Tables[0].Rows[intCount]["LastName"].ToString() + "\", \"ORCID\": \"" + ds.Tables[0].Rows[intCount]["ORCID"].ToString() + "\", \"ArticleAffiliation\": \"" + ds.Tables[0].Rows[intCount]["ArticleAffiliation"].ToString() + "\", \"RinggoldId\": \"" + ds.Tables[0].Rows[intCount]["RinggoldId"].ToString() + "\", \"RinggoldInstitution\": \"" + ds.Tables[0].Rows[intCount]["RinggoldInstitution"].ToString() + "\", \"IsCorrespondingAuthor\": " + blnCorrAuth + ", \"PrimaryEmail\": \"" + ds.Tables[0].Rows[intCount]["PrimaryEmail"].ToString() + "\", \"MailingAffiliation\": \"" + ds.Tables[0].Rows[intCount]["MailingAffiliation"].ToString() + "\", \"Address1\": \"" + ds.Tables[0].Rows[intCount]["Address1"].ToString() + "\", \"Address2\": \"" + ds.Tables[0].Rows[intCount]["Address2"].ToString() + "\", \"City\": \"" + ds.Tables[0].Rows[intCount]["City"].ToString() + "\", \"StateOrProvince\": \"" + ds.Tables[0].Rows[intCount]["StateOrProvince"].ToString() + "\", \"Country\": \"" + ds.Tables[0].Rows[intCount]["Country"].ToString() + "\", \"PostalCode\": \"" + ds.Tables[0].Rows[intCount]["Postalcode"].ToString() + "\", \"IsGroupAuthor\": " + blnGroupAuth + ", \"GroupAuthorName\": \"" + ds.Tables[0].Rows[intCount]["GroupAuthorName"].ToString() + "\", \"SequenceNo\": " + ds.Tables[0].Rows[intCount]["SequenceNo"].ToString() + "},";
                        }
                        strAuthorInfoCollection = strAuthorInfoCollection.Substring(0, strAuthorInfoCollection.Length - 1);
                    }
                    else {
                        return Json(new { dataComp = "Authors should not be empty." }, JsonRequestBehavior.AllowGet);
                    }



                    // Funder details collection
                    string strFundDetails = string.Empty;

                    if (strFundingInfo != "")
                    {
                        // Load the fund xml to object
                        try
                        {
                            XmlDocument objXML = new XmlDocument();
                            objXML.LoadXml(strFundingInfo);

                            XmlNodeList xnList = objXML.SelectNodes("//config/site/item[@isDefault='Yes']");
                            foreach (XmlNode xn in xnList)
                            {

                            }

                            strFundDetails = "\"OpenFunders\": [" + strFunderInfoCollection + "],";


                        }
                        catch (Exception)
                        {
                            strFundDetails = "";
                        }

                    }
                    else
                    { strFundDetails = "\"OpenFunders\": [],"; }

                    string jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + vArticleID + ", \"ArticleInfo\": { \"ArticleType\": \"" + A.AD.ArticleType + "\", \"OtherArticleType\": \"" + strOtherArticleType + "\", \"Title\": \"" + A.AD.ArticleTitle + "\", \"SubTitle\": \"" + strSubTitle + "\", \"Issue\": \"" + A.AD.iss + "\", \"IssueSection\": \"" + A.AD.IssType + "\", \"SpecialIssueOrCollectionsName\": \"" + strSpecialIssueOrCollectionsName + "\", \"IsExcludeFromOnlineFirst\": " + blnIsExcludeFromOnlineFirst + ", \"WordCount\": " + intWordCount + ", \"FigureCount\": " + intFigCount + ", \"TableCount\": " + intTableCount + ", \"ActualPages\": " + A.AD.intActualPages + ", \"IsLaTex\": " + blnIsLatex + ", \"IsContainsColour\": " + blnIsColor + ", \"IsColourFiguresRequirePayment\": " + blnIsColorPayment + ", \"IsSupplementalMaterial\": " + A.AD.SuppDataAvailable + ", \"ManuscriptNumber\": \"" + strManuscriptID + "\", \"ManuscriptType\": \"" + strManuscriptType + "\", \"AcceptanceDate\": \"" + A.AD.JN_AcceptedDate + "\", \"SubmittedDate\": \"" + strSubmittedDate + "\", \"RevisedSubmissionDate\": \"" + A.AD.JN_RevisedDate + "\"}," + strFundDetails + "\"Authors\": [" + strAuthorInfoCollection + "]}";

                    // To trigger API
                    HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, "EV056", vEventType, jsonContent);

                    //To update the event status in DB
                    string strStatus = Proc_UpdateAPIStatus(responseResult, A, "EV056", "", vEventType, jsonContent);

                    if (strStatus.ToLower().Contains("success"))
                    {
                        //strResult = DBProc.GetResultasString("Insert into JBM_ScheduleInfo(JBM_AutoID,AutoArtID,TaskName,Task_TAT,ResourceName,PlannedBy,Status) Values ('" + vJBMID + "','" + vAutoArtID + "','" + vTask + "','" + vTaskDays + "','" + vResourceName.Trim() + "','" + Session["EmpAutoId"].ToString() + "','Waiting for the event result')", Session["sConnSiteDB"].ToString());
                    }

                    return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);
                }


            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private string Proc_CountryStateValidation()
        {
            try
            {
                string strExcelPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Support\\SAGE-SMART Country and State Codes.xlsx");
                
                var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", strExcelPath);

                var adapter = new OleDbDataAdapter("SELECT * FROM [City]", connectionString);
                var ds = new DataSet();

                adapter.Fill(ds, "tblCountry");

                DataTable data = ds.Tables["tblCountry"];

                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private string Proc_ContentValidation(string SageArticleId, string ArticleType, string ArticleTitle, string blnIsExcludeFromOnlineFirst, int intWordCount, int intFigCount, int intTableCount, int intActualPages)
        {
            try
            {
                if (SageArticleId == "0")
                {
                    return "Article Id should not be zero.";
                }

                if (ArticleType == "")
                {
                    return "Article Type should not be empty.";
                }
                else if (ArticleType.Length > 255)
                {
                    return "Article Type must be maximum of 255 characters.";
                }

                if (ArticleTitle == "")
                {
                    return "Title should not be empty.";
                }
                

                if (blnIsExcludeFromOnlineFirst == "")
                {
                    return "Exclude From Online First should be either true or false.";
                }

                if (intWordCount.ToString() == "")
                {
                    return "Word count should be a valid number.";
                }

                if (intFigCount.ToString() == "")
                {
                    return "Figure count should be a valid number.";
                }

                if (intTableCount.ToString() == "")
                {
                    return "Table count should be a valid number.";
                }

                if (intActualPages.ToString() == "")
                {
                    return "Actual pages should be a valid number.";
                }

                return "";
            }
            catch (Exception)
            {
                return "";
            }

        }

        public ActionResult AAMFileOnlineEvent(string vJBMID, string vAutoArtID, string vArticleID, string vType)
        {
            try
            {
                string strFileCollection = "";
                using (var fileStream = System.IO.File.Open(@"D:\Testing\jcpt-2021-may-0140-20210812040235.zip", System.IO.FileMode.Open))
                {
                    var parentArchive = new ZipArchive(fileStream);

                    foreach (var e in parentArchive.Entries)
                    {
                        Console.WriteLine(e.Name  + "  " + e.Length);
                        strFileCollection += "{\"FileName\": \"" + e.Name + "\",\"FileSize\": \"" + e.Length + "\"},";
                        
                    }
                    strFileCollection = strFileCollection.Substring(0, strFileCollection.Length - 1);
                }

                string jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + vArticleID + ",\"LoadType\": \"" + vType + "\",\"Files\": [" + strFileCollection + "]}";

                RL.ArtDet A = new RL.ArtDet();
                A.AD.JAutoID = vJBMID;
                A.AD.AutoArtID = vAutoArtID;
                A.AD.ArticleID = vArticleID;
                string strEventID = "";
                if (vType == "AAMFileOutput")
                {
                    strEventID = "EV049";
                }
                else if (vType == "AAMFileOutput")
                {
                    strEventID = "EV049";

                }

                // To trigger API
                HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, strEventID, "ArticleLoadStatus", jsonContent);

                //To update the event status in DB
                string strStatus = Proc_UpdateAPIStatus(responseResult, A, strEventID, "", "ArticleLoadStatus", jsonContent);

                if (strStatus.ToLower().Contains("success"))
                {
                    //strResult = DBProc.GetResultasString("Insert into JBM_ScheduleInfo(JBM_AutoID,AutoArtID,TaskName,Task_TAT,ResourceName,PlannedBy,Status) Values ('" + vJBMID + "','" + vAutoArtID + "','" + vTask + "','" + vTaskDays + "','" + vResourceName.Trim() + "','" + Session["EmpAutoId"].ToString() + "','Waiting for the event result')", Session["sConnSiteDB"].ToString());
                }

                return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);
                

            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult addNewTaskDetails(string vJBMID, string vAutoArtID, string vTask, string vTaskDays, string vResourceName, string vArticleID)
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select TaskName from JBM_ScheduleInfo WHERE TaskName='" + vTask.Trim() + "' and JBM_AutoID='" + vJBMID.Trim() + "' and AutoArtID='" + vAutoArtID.Trim() + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count == 0)
                {
                    string strResult = string.Empty;
                    string strTaskCollection = string.Empty;
                    strTaskCollection = @"{""Task"": """ + vTask.ToString() + @""",""TaskDays"": " + vTaskDays.ToString() + @",""ResourceName"": """ + vResourceName.Trim().ToString() + @""",""OperationType"": ""Add""}";
                    RL.ArtDet A = new RL.ArtDet();
                    A.AD.JAutoID = vJBMID;
                    A.AD.AutoArtID = vAutoArtID;
                    A.AD.ArticleID = vArticleID;
                    A.AD.Stage = vTask.ToString();
                    A.AD.Arr[0] = vTaskDays.ToString();
                    A.AD.UsrName = vResourceName.Trim().ToString();

                    string jsonContent = jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + A.AD.ArticleID + ",\"Tasks\": [" + strTaskCollection + "]}";

                    // To trigger API
                    HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, "EV052", "UpdateArticleSchedule", jsonContent);

                    //To update the event status in DB
                    string strStatus = Proc_UpdateAPIStatus(responseResult, A, "EV052", "AddTask", "UpdateArticleSchedule", strTaskCollection);

                    if (strStatus.ToLower().Contains("added") || strStatus.ToLower().Contains("success"))
                    {
                        //strResult = DBProc.GetResultasString("Insert into JBM_ScheduleInfo(JBM_AutoID,AutoArtID,TaskName,Task_TAT,ResourceName,PlannedBy,Status) Values ('" + vJBMID + "','" + vAutoArtID + "','" + vTask + "','" + vTaskDays + "','" + vResourceName.Trim() + "','" + Session["EmpAutoId"].ToString() + "','Waiting for the event result')", Session["sConnSiteDB"].ToString());
                    }

                    return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);
                   
                }
                else
                {
                    return Json(new { dataComp = "Task already exists for this article." }, JsonRequestBehavior.AllowGet);
                }
              
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message}, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult removeTaskDetails(string ItemColloc)
        {
            try
            {
                string strAutoIDColl = string.Empty;
                List<string> taskIds = JsonConvert.DeserializeObject<List<string>>(ItemColloc);
                if (taskIds.Count > 0)
                {
                    string strTaskCollection = string.Empty;
                    string vAutoID = string.Empty; string vTask = string.Empty; string vTaskDays = string.Empty; string vResourceName = string.Empty; string vJBMAutoID = string.Empty; string vAutoArtID = string.Empty; string vArticleID = string.Empty;

                    for (int i = 0; i < taskIds.Count; i++)
                    {
                        vAutoID = string.Empty; vTask = string.Empty; vTaskDays = string.Empty; vResourceName = string.Empty; vJBMAutoID = string.Empty; vAutoArtID = string.Empty; vArticleID = string.Empty;
                        vAutoID = taskIds[i].Split('|')[0];
                        vTask = taskIds[i].Split('|')[1];
                        vTaskDays = taskIds[i].Split('|')[2];
                        vResourceName = taskIds[i].Split('|')[3];
                        vJBMAutoID = taskIds[i].Split('|')[4];
                        vAutoArtID = taskIds[i].Split('|')[5];
                        vArticleID = taskIds[i].Split('|')[6];

                        strAutoIDColl += vAutoID + ",";

                        strTaskCollection += @"{""Task"": """ + vTask.ToString() + @""",""TaskDays"": " + vTaskDays.ToString() + @",""ResourceName"": """ + vResourceName.Trim().ToString() + @""",""OperationType"": ""Remove""},";
                    }

                    if (strAutoIDColl[strAutoIDColl.Length - 1] == ',')
                    {
                        strAutoIDColl = strAutoIDColl.Substring(0, strAutoIDColl.Length - 1);
                    }

                    if (strTaskCollection[strTaskCollection.Length - 1] == ',')
                    {
                        strTaskCollection = strTaskCollection.Substring(0, strTaskCollection.Length - 1);
                    }


                    RL.ArtDet A = new RL.ArtDet();
                    A.AD.JAutoID = vJBMAutoID;
                    A.AD.AutoArtID = vAutoArtID;
                    A.AD.ArticleID = vArticleID;
                    string jsonContent = jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + A.AD.ArticleID + ",\"Tasks\": [" + strTaskCollection + "]}";

                    // To trigger API
                    HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, "EV052", "UpdateArticleSchedule", jsonContent);

                    //To update the event status in DB
                    string strStatus = Proc_UpdateAPIStatus(responseResult, A, "EV052", strAutoIDColl, "UpdateArticleSchedule", strTaskCollection);
                   
                    if (strStatus.ToLower().Contains("updated") || strStatus.ToLower().Contains("success") || strStatus.ToLower().Contains("removed"))
                    {
                        // Remove the task from end
                        string strResult = DBProc.GetResultasString("DELETE FROM JBM_ScheduleInfo WHERE AutoID in (" + strAutoIDColl + ")", Session["sConnSiteDB"].ToString());
                    }

                    return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult updateTaskDetails(string ItemColloc)
        {
            try
            {
                List<string> taskIds = JsonConvert.DeserializeObject<List<string>>(ItemColloc);
                if (taskIds.Count > 0)
                {
                    string strAutoIDColl = string.Empty;
                    string strTaskCollection = string.Empty;
                    string vAutoID = string.Empty; string vTask = string.Empty; string vTaskDays = string.Empty; string vResourceName = string.Empty; string vJBMAutoID = string.Empty; string vAutoArtID = string.Empty; string vArticleID = string.Empty;

                    for (int i = 0; i < taskIds.Count; i++)
                    {
                        vAutoID = string.Empty; vTask = string.Empty; vTaskDays = string.Empty; vResourceName = string.Empty; vJBMAutoID = string.Empty; vAutoArtID = string.Empty; vArticleID = string.Empty;
                        vAutoID = taskIds[i].Split('|')[0];
                        vTask = taskIds[i].Split('|')[1];
                        vTaskDays = taskIds[i].Split('|')[2];
                        vResourceName = taskIds[i].Split('|')[3];
                        vJBMAutoID = taskIds[i].Split('|')[4];
                        vAutoArtID = taskIds[i].Split('|')[5];
                        vArticleID = taskIds[i].Split('|')[6];

                        strAutoIDColl += vAutoID + ",";

                        strTaskCollection += @"{""Task"": """ + vTask.ToString() + @""",""TaskDays"": "+ vTaskDays.ToString() + @",""ResourceName"": """ + vResourceName.ToString() + @""",""OperationType"": ""Update""},";
                    }

                    if (strTaskCollection[strTaskCollection.Length - 1] == ',')
                    {
                        strTaskCollection = strTaskCollection.Substring(0, strTaskCollection.Length - 1);
                    }

                    if (strAutoIDColl[strAutoIDColl.Length - 1] == ',')
                    {
                        strAutoIDColl = strAutoIDColl.Substring(0, strAutoIDColl.Length - 1);
                    }

                    RL.ArtDet A = new RL.ArtDet();
                    A.AD.JAutoID = vJBMAutoID;
                    A.AD.AutoArtID = vAutoArtID;
                    A.AD.ArticleID = vArticleID;
                    string jsonContent = jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + A.AD.ArticleID + ",\"Tasks\": [" + strTaskCollection + "]}";

                    // To trigger API
                    HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, "EV052", "UpdateArticleSchedule", jsonContent);

                    //To update the event status in DB
                    string strStatus = Proc_UpdateAPIStatus(responseResult, A, "EV052", strAutoIDColl, "UpdateArticleSchedule", strTaskCollection);

                    if (strStatus.Contains("updated"))
                    {
                        for (int i = 0; i < taskIds.Count; i++)
                        {
                            vAutoID = string.Empty; vTask = string.Empty; vTaskDays = string.Empty; vResourceName = string.Empty; vJBMAutoID = string.Empty; vAutoArtID = string.Empty; vArticleID = string.Empty;
                            vAutoID = taskIds[i].Split('|')[0];
                            vTask = taskIds[i].Split('|')[1];
                            vTaskDays = taskIds[i].Split('|')[2];
                            vResourceName = taskIds[i].Split('|')[3];
                            vJBMAutoID = taskIds[i].Split('|')[4];
                            vAutoArtID = taskIds[i].Split('|')[5];
                            vArticleID = taskIds[i].Split('|')[6];

                            //Update in the DB
                            string strRes = DBProc.GetResultasString("UPDATE JBM_ScheduleInfo SET Task_TAT='" + vTaskDays.ToString() + @"',ResourceName='" + vResourceName.ToString() + "', Status='" + strStatus.Replace("'", "''") + "' WHERE AutoID='" + vAutoID + "'", Session["sConnSiteDB"].ToString());

                        }

                    }

                    return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);


                }
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult completeTaskDetails(string ItemColloc)
        {
            try
            {
                List<string> taskIds = JsonConvert.DeserializeObject<List<string>>(ItemColloc);
                if (taskIds.Count > 0)
                {
                    
                    string strTaskCollection = string.Empty;
                    string vAutoID = string.Empty; string vTask = string.Empty; string vTaskDays = string.Empty; string vResourceName = string.Empty; string vJBMAutoID = string.Empty; string vAutoArtID = string.Empty; string vArticleID = string.Empty;

                    for (int i = 0; i < taskIds.Count; i++)
                    {
                        vAutoID = string.Empty; vTask = string.Empty; vTaskDays = string.Empty; vResourceName = string.Empty; vJBMAutoID = string.Empty; vAutoArtID = string.Empty; vArticleID = string.Empty;
                        vAutoID = taskIds[i].Split('|')[0];
                        vTask = taskIds[i].Split('|')[1];
                        vTaskDays = taskIds[i].Split('|')[2];
                        vResourceName = taskIds[i].Split('|')[3];
                        vJBMAutoID = taskIds[i].Split('|')[4];
                        vAutoArtID = taskIds[i].Split('|')[5];
                        vArticleID = taskIds[i].Split('|')[6];

                        break;                      
                    }

                    RL.ArtDet A = new RL.ArtDet();
                    A.AD.JAutoID = vJBMAutoID;
                    A.AD.AutoArtID = vAutoArtID;
                    A.AD.ArticleID = vArticleID;
                    string jsonContent = jsonContent = "{\"VendorCode\": \"V_KGL\",\"SageArticleId\": " + A.AD.ArticleID + ",\"Task\": \"" + vTask + "\", \"ResourceName\": \"" + vResourceName + "\"}";

                    // To trigger API
                    HttpResponseMessage responseResult = Proc_InstantSageAPITrigger(A, "EV053", "CompleteTask", jsonContent);

                    //To update the event status in DB
                    string strStatus = Proc_UpdateAPIStatus(responseResult, A, "EV053", vAutoID, "CompleteTask", jsonContent);

                    if (strStatus.Contains("success") || strStatus.Contains("completed"))
                    {
                        //Update in the DB
                        string strRes = DBProc.GetResultasString("UPDATE JBM_ScheduleInfo SET ResourceName='" + vResourceName.ToString() + "', Status='" + strStatus.Replace("'", "''") + "' WHERE AutoID='" + vAutoID + "'", Session["sConnSiteDB"].ToString());

                    }

                    return Json(new { dataComp = strStatus }, JsonRequestBehavior.AllowGet);


                }
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public string Proc_UpdateAPIStatus(HttpResponseMessage responseResult, ReferenceLibrary.ArtDet A, string strEventID, string strAutoIDColl, string strLoadType, string strEventJson)
        {
            try
            {

                string strResponse;
                string strResult;

                //string ss = "{\"Status\":\"Failed\",\"StatusCode\":400,\"Message\":\"Start date should be current date or future date.\",\"Data\":null}";
                //var jsonResul = JsonConvert.DeserializeObject<Dictionary<string, string>>(ss);
                //string name = (string)jsonResul["Message"];

                strResponse = responseResult.Content.ReadAsStringAsync().Result;

                DataSet dsJson = new DataSet();
                XmlDocument xd = new XmlDocument();

                string jsonString = "";

                jsonString = "{ \"rootNode\": {" + strResponse.Trim().TrimStart('{').TrimEnd('}') + "} }";
                xd = (XmlDocument)Newtonsoft.Json.JsonConvert.DeserializeXmlNode(jsonString);

                string strMessage = "";
                string strStatus = "";
                dsJson.ReadXml(new XmlNodeReader(xd));
                if (dsJson != null)
                {
                    if (dsJson.Tables["rootNode"].Rows.Count > 0)
                    {
                        foreach (DataRow row in dsJson.Tables["rootNode"].Rows)
                        {
                            strMessage = Convert.ToString(row["Message"]);
                            strStatus = Convert.ToString(row["Status"]);
                        }
                    }

                }


                if (responseResult.IsSuccessStatusCode)
                {
                    strResponse = responseResult.Content.ReadAsStringAsync().Result;
                    strResponse = strResponse.Replace(@"\", "");
                    strResponse = strResponse.Trim('"');

                    var jsonResult = JsonConvert.DeserializeObject<Dictionary<string, string>>(strResponse);
                    if (jsonResult.ContainsKey("Status"))
                    {
                      
                        if (strStatus.ToString().ToLower() == "failed")
                        {
                            if (strAutoIDColl != "AddTask")
                            {
                                strResult = DBProc.GetResultasString("UPDATE JBM_ScheduleInfo SET Status='" + strMessage.Replace("'", "''") + "' WHERE AutoID in (" + strAutoIDColl + ")", Session["sConnSiteDB"].ToString());
                            }
                            strResult = DBProc.GetResultasString("INSERT INTO SG_ExternalEventAccess (JBM_AutoID,ArticleID,AutoArtID,ExternalEventID,ExternalEventStatus,ExternalEventTriggerd,StatusDescription,MaxTry,Status,EventCallback) VALUES ('" + A.AD.JAutoID + "','" + A.AD.ArticleID + "','" + A.AD.AutoArtID + "','" + strEventID + "','Failure',GETDATE(),'" + strResponse + "',0,0,'" + strEventJson + "')", Session["sConnSiteDB"].ToString());
                            strResult = DBProc.GetResultasString("Insert into JBM_ProdAccess (CustAcc, AutoArtID, JBM_AutoID, EmpAutoID, AccTime, Process, AccPage, Descript, ShortDescript) values ('SG','" + A.AD.AutoArtID + "','" + A.AD.JAutoID + "','E000001',GETDATE(),'Event Failure','Sage API Call','" + strMessage.Replace("'", "''") + "', '" + strLoadType + "')", Session["sConnSiteDB"].ToString());

                            return strMessage;
                        }
                        else
                        {
                            if (strAutoIDColl == "AddTask")
                            {
                                strResult = DBProc.GetResultasString("Insert Into JBM_ScheduleInfo(JBM_AutoID,AutoArtID,TaskName,Task_TAT,ResourceName,PlannedBy,Status) Values ('" + A.AD.JAutoID + "','" + A.AD.AutoArtID + "','" + A.AD.Stage + "','" + A.AD.Arr[0].ToString() + "','" + A.AD.UsrName.ToString() + "','" + Session["EmpAutoId"].ToString() + "','" + strMessage.Replace("'", "''") + "')", Session["sConnSiteDB"].ToString());
                            }
                            else
                            {
                                strResult = DBProc.GetResultasString("UPDATE JBM_ScheduleInfo SET Status='" + strMessage.Replace("'", "''") + "' WHERE AutoID in (" + strAutoIDColl + ")", Session["sConnSiteDB"].ToString());
                            }
                            strResult = DBProc.GetResultasString("INSERT INTO SG_ExternalEventAccess (JBM_AutoID,ArticleID,AutoArtID,ExternalEventID,ExternalEventStatus,ExternalEventTriggerd,StatusDescription,MaxTry,Status,EventCallback) VALUES ('" + A.AD.JAutoID + "','" + A.AD.ArticleID + "','" + A.AD.AutoArtID + "','" + strEventID + "','Success',GETDATE(),'" + strResponse + "',1,1,'" + strEventJson + "')", Session["sConnSiteDB"].ToString());
                            strResult = DBProc.GetResultasString("Insert into JBM_ProdAccess (CustAcc, AutoArtID, JBM_AutoID, EmpAutoID, AccTime, Process, AccPage, Descript, ShortDescript) values ('SG','" + A.AD.AutoArtID + "','" + A.AD.JAutoID + "','E000001',GETDATE(),'Event Success','Sage API Call','" + strMessage.Replace("'", "''") + "', '" + strLoadType + "')", Session["sConnSiteDB"].ToString());

                            return strMessage;
                        }
                    }
                }
                else
                {
                    strResponse = responseResult.Content.ReadAsStringAsync().Result;
                    if (strAutoIDColl != "AddTask")
                    {
                        strResult = DBProc.GetResultasString("UPDATE JBM_ScheduleInfo SET Status='" + strMessage.Replace("'", "''") + "' WHERE AutoID in (" + strAutoIDColl + ")", Session["sConnSiteDB"].ToString());
                    }
                    strResult = DBProc.GetResultasString("INSERT INTO SG_ExternalEventAccess (JBM_AutoID,ArticleID,AutoArtID,ExternalEventID,ExternalEventStatus,ExternalEventTriggerd,StatusDescription,MaxTry,Status,EventCallback) VALUES ('" + A.AD.JAutoID + "','" + A.AD.ArticleID + "','" + A.AD.AutoArtID + "','" + strEventID + "','Failure',GETDATE(),'" + strResponse + "',0,0,'" + strEventJson + "')", Session["sConnSiteDB"].ToString());
                    strResult = DBProc.GetResultasString("Insert into JBM_ProdAccess (CustAcc, AutoArtID, JBM_AutoID, EmpAutoID, AccTime, Process, AccPage, Descript, ShortDescript) values ('SG','" + A.AD.AutoArtID + "','" + A.AD.JAutoID + "','E000001',GETDATE(),'Event Failure','Sage API Call','" + strMessage.Replace("'", "''") + "', '" + strLoadType + "')", Session["sConnSiteDB"].ToString());
                    return strMessage;
                }
                return "Success";
            }
            catch (Exception)
            {
                return "Failed";
            }
        }


        /// <summary>
        /// To trigger the SAGE API
        /// </summary>
        /// <param name="EventContent"></param>
        /// <param name="strEndPoint"></param>
        /// <param name="A"></param>
        /// <param name="strEventID"></param>
        /// <param name="strLoadType"></param>
        /// <param name="strEventJson"></param>
        /// <returns></returns>
        public HttpResponseMessage Proc_InstantSageAPITrigger(ReferenceLibrary.ArtDet A, string strEventID, string strLoadType, string strEventJson)
        {
            try
            {

                HttpClient client = new HttpClient();

                //specify to use TLS 1.2 as default connection
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
               
                client.DefaultRequestHeaders.Add("Event-Environment", GlobalVariables.strEnvironment.Trim().ToString());  // LIVE or UAT
                client.DefaultRequestHeaders.Add("Event-Type", strLoadType);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "kgluatuser", "526daba8bda630f8e5d991ad78275c5e"))));

                var jsonData = new StringContent(strEventJson, Encoding.UTF8, "application/json");

                HttpResponseMessage response = client.PostAsync("https://smarttrack.kwglobal.com/sage.api.kglind-" + GlobalVariables.strEnvironment.Trim().ToString().ToLower() + "/v1/sageapi" , jsonData).Result;
                return response;

                ////HttpClient client = new HttpClient();

                //////specify to use TLS 1.2 as default connection
                ////ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ////System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

                ////var plainTextBytes = System.Text.Encoding.UTF8.GetBytes("kgluatuser:526daba8bda630f8e5d991ad78275c5e");
                ////string val = System.Convert.ToBase64String(plainTextBytes);
                ////client.DefaultRequestHeaders.Add("Authorization", "Basic " + val);
                ////client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //////client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "kgluatuser", "526daba8bda630f8e5d991ad78275c5e"))));

                ////var jsonData = new StringContent(strEventJson, Encoding.UTF8, "application/json");

                ////HttpResponseMessage response = client.PostAsync("https://journalsuat.sageapps.com/VendorApi/v1/" + strLoadType, jsonData).Result;
                ////return response;

            }
            catch (Exception)
            {
                return null;
            }
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

    }
}