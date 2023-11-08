using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using SmartTrack.Helper;
using System.Xml;
using RL = ReferenceLibrary;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace SmartTrack.Controllers
{
    public class CommonController : Controller
    {
        // GET: Common

        private bool blnIEEEview = false;
        DataProc DBProc = new DataProc(); // Data store/retrive DB

        public ActionResult CommonControls(string controls = "")
        {
            string strSiteID = clsInit.gStrLocationID;
            ViewBag.controls = controls;

            if (Session["gJwAccItm"].ToString().Contains("|IEV|") & Session["sCustAcc"].ToString() == "MG")
            {
                blnIEEEview = true;
            }

            ViewBag.TeamList = LoadTeamDetails();

            if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupJrnls001 | Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
            {
                ViewBag.selectedTeamID = Session["sCustTeamID"].ToString().Replace("|", "");
            }
            else
            {
                ViewBag.selectedTeamID = "";
            }
            return PartialView();
        }

        private List<SelectListItem> LoadTeamDetails()
        {
            if (Session["sCustGroup"].ToString() == "BK" | Session["sCustGroup"].ToString() == "MG")
            {
                return bindToSelectListItem(DBProc.GetResultasDataTbl("select TeamID, Description from " + Init_Tables.gTblJBMCustTeamID + " where custType='" + Session["sCustAcc"].ToString() + "' " + Session["gTeamID"].ToString().Replace("|", "") + " order by teamid desc", Session["sConnSiteDB"].ToString()), "Description", "TeamID");
            }
            else
            {
                return bindToSelectListItem(DBProc.GetResultasDataTbl("select TeamID, Description from " + Init_Tables.gTblJBMCustTeamID + " where custType='" + Session["sCustAcc"].ToString() + "' order by teamid desc", Session["sConnSiteDB"].ToString()), "Description", "TeamID");
            }
        }

        public List<SelectListItem> bindToSelectListItem(DataTable dt, string strTextField, string strValueField)
        {
            List<SelectListItem> list = new List<SelectListItem>();

            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    list.Add(new SelectListItem()
                    {
                        Text = dr[strTextField].ToString(),
                        Value = dr[strValueField].ToString()
                    });
                }
            }
            return list;
        }
        public List<SelectListItem> getSubTeamList(string TeamID)
        {
            List<SelectListItem> SubTeamList = new List<SelectListItem>();
            if (Session["sCustAcc"].ToString().ToLower() != "jw")
            {
                SubTeamList = bindToSelectListItem(DBProc.GetResultasDataTbl("select SubTeamID,SubTeamName from " + Init_Tables.gTblJBM_SubTeam + " where TeamID='" + TeamID + "' order by SubTeamID", Session["sConnSiteDB"].ToString()), "SubTeamName", "SubTeamID");
            }

            return SubTeamList;
        }

        public ActionResult getCustomerDetails(string TeamID)
        {
            try
            {
                string strTeamIDQry = TeamID != null ? " and JBM_TeamID=" + TeamID + "" : "";

                if (TeamID != "")
                {
                    List<SelectListItem> CustomerList = bindToSelectListItem(DBProc.GetResultasDataTbl("select CustID,CustSN from " + Init_Tables.gTblCustomerMaster + " where CustType='" + Session["sCustAcc"].ToString() + "' " + strTeamIDQry + " Order by CustSN", Session["sConnSiteDB"].ToString()), "CustSN", "CustID");

                    return Json(CustomerList, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json("");
                }
            }
            catch(Exception ex)
            {
                return Json("");
            }      
        }

        public ActionResult getJournalDetails(string TeamID, string SubTeamID, string CustID)
        {
            List<SelectListItem> JournalList = new List<SelectListItem>();
            string strCustQry = "";
            string strComposition = "";
            string strEmp = "";

            if (Session["DeptCode"].ToString() == Init_Department.gDcIm & Session["gJwAccItm"].ToString().IndexOf("|IMA|") == -1 & Session["sCustAcc"].ToString().ToLower() != "jw")
            {
                strEmp = " and KGL_IM='" + Session["EmpAutoId"].ToString() + "'";
            }

            if (CustID != "" & CustID != null)
            {
                strCustQry = " and CustID='" + CustID + "' ";
            }

            if (Init_Department.gDcGraphics == "90" & Session["sCustAcc"].ToString().ToLower() == "jw")
            {
                strComposition = " and Composition=1 ";
            }

            string strSQL = "Select JBM_AutoID, JBM_Intrnl from " + Init_Tables.gTblJrnlInfo + " where (JBM_AutoID LIKE '" + Session["sCustAcc"].ToString() + "%') and JBM_Disabled='0' " + strComposition + strCustQry + " order by JBM_Intrnl";

            if (TeamID != "" | TeamID == null)
            {
                if (Session["sCustGroup"].ToString() == clsInit.gstrCustGroupJrnls001 | Session["sCustGroup"].ToString() == clsInit.gstrCustGroupSplAcc002)
                {
                    JournalList = bindToSelectListItem(DBProc.GetResultasDataTbl(strSQL, Session["sConnSiteDB"].ToString()), "JBM_Intrnl", "JBM_AutoID");

                    if (Session["sCustAcc"].ToString().ToLower() == "jw")
                    {
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-C", Value = "TEMC" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-B", Value = "TEMB" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-A", Value = "TEMA" });
                        JournalList.Insert(0, new SelectListItem() { Text = "FL-OCE", Value = "USCE" });
                        JournalList.Insert(0, new SelectListItem() { Text = "FL-ICE", Value = "FLCE" });
                        JournalList.Insert(0, new SelectListItem() { Text = "OTS", Value = "OTS" });
                        JournalList.Insert(0, new SelectListItem() { Text = "IHTS", Value = "IHTS" });
                        JournalList.Insert(0, new SelectListItem() { Text = "US Titles", Value = "US" });
                        JournalList.Insert(0, new SelectListItem() { Text = "UK Titles", Value = "UK" });
                        JournalList.Insert(0, new SelectListItem() { Text = "5 Days TAT Titles", Value = "5D" });
                        JournalList.Insert(0, new SelectListItem() { Text = "Direct Wiley Tiltles", Value = "DW" });
                        JournalList.Insert(0, new SelectListItem() { Text = "IM Tiltles", Value = "IM" });
                        JournalList.Insert(0, new SelectListItem() { Text = "AGU Tiltles", Value = "AGU" });
                    }
                    else if (Session["sCustAcc"].ToString() == "OP")
                    {
                        JournalList.Insert(0, new SelectListItem() { Text = "US Titles", Value = "US" });
                        JournalList.Insert(0, new SelectListItem() { Text = "UK Titles", Value = "UK" });
                    }
                    else if (Session["sCustAcc"].ToString().ToLower() == "tf")
                    {
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-F", Value = "TEMF" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-E", Value = "TEME" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-D", Value = "TEMD" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-C", Value = "TEMC" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-B", Value = "TEMB" });
                        JournalList.Insert(0, new SelectListItem() { Text = "TEAM-A", Value = "TEMA" });
                        JournalList.Insert(0, new SelectListItem() { Text = "US Titles", Value = "US" });
                        JournalList.Insert(0, new SelectListItem() { Text = "UK Titles", Value = "UK" });
                    }
                }
                else
                {

                    string SubTeamFilter = "";

                    if (SubTeamID != "" & SubTeamID != null)
                    {
                        SubTeamFilter = " and (JBM_SubTeam= " + SubTeamID + " or JBM_SubTeam is null) ";
                    }

                    strSQL = "Select JBM_AutoID, JBM_ID from " + Init_Tables.gTblJrnlInfo + " where (JBM_AutoID LIKE '" + Session["sCustAcc"].ToString() + "%') and JBM_Disabled='0' and JBM_TeamID = " + TeamID + SubTeamFilter + strCustQry + " order by JBM_Intrnl";

                    JournalList = bindToSelectListItem(DBProc.GetResultasDataTbl(strSQL, Session["sConnSiteDB"].ToString()), "JBM_ID", "JBM_AutoID");
                }

                return Json(JournalList, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json("", JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult Index(string CustAcc, string EmpID, string SiteID, string URL)
        {
            // To clear all connection pools
            SqlConnection.ClearAllPools();
            GlobalVariables.strSiteCode = null;
            GlobalVariables.strSiteLocation = null;
            GlobalVariables.strEnvironment = null;

            if (GlobalVariables.strSiteCode == null || GlobalVariables.strSiteLocation == null)
            {
                String strURL;
                strURL = System.Web.HttpContext.Current.Request.Url.Host;

                string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                XmlDocument xml = new XmlDocument();
                xml.Load(strPath);

                XmlNodeList xnList = xml.SelectNodes("//config/site/item[@isDefault='Yes']");
                foreach (XmlNode xn in xnList)
                {
                    Console.WriteLine(xn.InnerText);
                    var nameAttribute = xn.Attributes["value"];
                    if (nameAttribute != null)
                        GlobalVariables.strSiteCode = nameAttribute.Value.ToString();

                    var nameAttrID = xn.Attributes["id"];
                    if (nameAttrID != null)
                        GlobalVariables.strSiteID = nameAttrID.Value.ToString();

                    var nameAttrEnviron = xn.Attributes["environment"];
                    if (nameAttrEnviron != null)
                        GlobalVariables.strEnvironment = nameAttrEnviron.Value.ToString();

                    GlobalVariables.strSiteLocation = xn.InnerText;
                    break;
                }
            }

            //smartConfig();
            smartConfig(CustAcc.ToUpper() == "TF" ? "L0004" : SiteID, GlobalVariables.strEnvironment.Trim().ToString());
            Session["newSmartTrack"] = "No";
            Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
            Session["sCustAcc"] = CustAcc;
            Session["returnURL"] = URL;

            string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno,a.NASAccessCmd from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + EmpID + "' or a.EmpMailid = '" + EmpID + "') and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

            InitializeSession(ds, CustAcc);
            smartConfig(CustAcc.ToUpper() == "TF" ? "L0004" : SiteID, GlobalVariables.strEnvironment.Trim().ToString());
            Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();

            return RedirectToAction(URL);
        }

        private void smartConfig(string SiteID = null, string Environment = null)
        {
            // To clear all connection pools
            SqlConnection.ClearAllPools();
            GlobalVariables.strSiteCode = null;
            GlobalVariables.strSiteLocation = null;
            GlobalVariables.strEnvironment = null;

            if (GlobalVariables.strSiteCode == null || GlobalVariables.strSiteLocation == null)
            {
                String strURL;
                strURL = System.Web.HttpContext.Current.Request.Url.Host;

                string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                XmlDocument xml = new XmlDocument();
                xml.Load(strPath);

                XmlNodeList xnList = null;

                if (SiteID == null)
                {
                    xnList = xml.SelectNodes("//config/site/item[@isDefault='Yes']");
                }
                else
                {
                    xnList = xml.SelectNodes("//config/site/item[@environment='" + Environment + "'][@id='" + SiteID + "']");
                }

                foreach (XmlNode xn in xnList)
                {
                    Console.WriteLine(xn.InnerText);
                    var nameAttribute = xn.Attributes["value"];
                    if (nameAttribute != null)
                        GlobalVariables.strSiteCode = nameAttribute.Value.ToString();

                    var nameAttrID = xn.Attributes["id"];
                    if (nameAttrID != null)
                        GlobalVariables.strSiteID = nameAttrID.Value.ToString();

                    var nameAttrEnviron = xn.Attributes["environment"];
                    if (nameAttrEnviron != null)
                        GlobalVariables.strEnvironment = nameAttrEnviron.Value.ToString();

                    GlobalVariables.strSiteLocation = xn.InnerText;
                    break;
                }

            }
        }

        public string InitializeSession(DataSet ds, string custAcc = "")
        {
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    //TempData["dept"] = ds.Tables[0].Rows[0][6].ToString();
                    Session["UserID"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
                    Session["EmpPass"] = ds.Tables[0].Rows[0]["EmpPass"].ToString();
                    Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
                    Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();
                    Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                    Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                    Session["DeptAcc"] = ds.Tables[0].Rows[0]["DeptAccess"].ToString();
                    Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();
                    Session["gJwAccItm"] = ds.Tables[0].Rows[0]["JwAccessItm"].ToString();
                    Session["gTeamID"] = ds.Tables[0].Rows[0]["TeamID"].ToString();
                    Session["AccessRights"] = "";
                    Session["CustomerSN"] = "";
                    Session["CustName"] = "";
                    Session["strHomeURL"] = "";
                    Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
                    Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
                    Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();
                    Session["NASAccessCmd"] = ds.Tables[0].Rows[0]["NASAccessCmd"].ToString();

                    string CustAccess = ds.Tables[0].Rows[0]["CustAccess"].ToString();
                    if (CustAccess != "")
                    {
                        CustAccess = custAcc == "" ? CustAccess.Substring(1, CustAccess.IndexOf("|", 1) - 1) : custAcc;
                        string strAccType = "Select ATD.AccountType, ATD.CustGroup, ATD.JBM_TeamID, ATD.Job_ID, (Select (case when ATD.RootID IS NULL then '' else R.RootPath end) from " + Init_Tables.gTblRootDirectory + " R where R.RootID=ATD.RootID) as RootDirPath, (Select (case when ATD.InwardDir IS NULL then '' else R.RootPath end) from " + Init_Tables.gTblRootDirectory + " R where R.RootID=ATD.InwardDir) as InwardDirPath, ATD.CustID, ATD.MLDir, ATD.GraphicsDirPath, ATD.MyPetInDir, ATD.MyPetOutDir, ATD.InwardFPDirPath, ATD.InwardRevDirPath, ATD.InwardIssueDirPath, ATD.DisplayCustIDPD, ATD.CreateDirectory, ATD.ImFpDirPath, ATD.ImRevDirPath, ATD.ImIssueDirPath, ATD.CeFpDirPath, ATD.CeRevDirPath, ATD.CeIssueDirPath, ATD.ReFpDirPath, ATD.ReRevDirPath, ATD.ReIssueDirPath, ATD.MLFpDirPath, ATD.MLRevDirPath, ATD.MLIssueDirPath, ATD.PagFpDirPath, ATD.PagRevDirPath, ATD.PagIssueDirPath, ATD.PrFpDirPath, ATD.PrRevDirPath, ATD.PrIssueDirPath, ATD.DispatchFpDirPath, ATD.DispatchRevDirPath, ATD.DispatchIssueDirPath, ATD.EproofCwpDir, ATD.DispatchVolDirPath, ATD.ErrorDirPath, ATD.CommentEnableDir, ATD.EproofNormal, ATD.DispatchUploadIn, ATD.DispatchSproofDirPath, ATD.SiteId, ATD.GraphicsOnlineDirPath, ATD.GraphicsWebDirPath, ATD.MarkPDFDirPath, ATD.CWP_CreatorEmail, ATD.IssueManagement, ATD.WorkingFolderDirPath, ATD.IceDirPath, ATD.DispatchHTML, ATD.ConvFpDirPath  from " + Init_Tables.gTblAccountTypeDesc + " ATD where ATD.CustAccess='" + CustAccess + "'";
                        DataSet dsAccType = DBProc.GetResultasDataSet(strAccType, Session["sConnSiteDB"].ToString()); //dbConnSmartTrack

                        if (dsAccType.Tables[0].Rows.Count > 0)
                        {
                            Session["sCustGroup"] = dsAccType.Tables[0].Rows[0]["CustGroup"];
                            Session["sCustTeamID"] = dsAccType.Tables[0].Rows[0]["JBM_TeamID"];
                            Session["InwardDirPath"] = dsAccType.Tables[0].Rows[0]["InwardDirPath"];
                            Session["AccountType"] = dsAccType.Tables[0].Rows[0]["AccountType"].ToString();
                        }
                    }

                }
                return "Success";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
        public ActionResult FLConfirm(string info)
        {
            try
            {
                RL.ArtDet A = new RL.ArtDet();
                SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
                info = HttpUtility.UrlDecode(info);
                info = info.Replace(" ", "+");
                string strreturnURL = objDS.DecryptData(info);

                string[] strSplit = strreturnURL.Split('|');

                A.AD.DBLoc = strSplit[0].ToString();
                A.AD.CustAccess = strSplit[1].ToString();
                A.AD.UsrLogin = strSplit[2].ToString();
                A.AD.AutoArtID = strSplit[4].ToString().Replace(",", "','");
                A.AD.Stage = strSplit[3].ToString();
                A.AD.DeptCode = strSplit[3].ToString();
                A.AD.JrnlSiteID = "L0004";
                string strConnectionSite = "dbConnSmartBG-LIVE"; //"dbConnSmart" + A.AD.DBLoc + "-LIVE";
                //string strConnectionSite = "dbConnSmartCH-LIVE";

                string[] strAutoArtIDs = A.AD.AutoArtID.Split(',');
                DataTable ds = new DataTable();
                DataTable dsFinal = new DataTable();
                //ds = DBProc.GetResultasDataSet("Select ai.ChapterID,ai.JBM_AutoID,ai.AutoArtID,al.AllocatedTo,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName],al.AcceptanceYN,CONVERT(varchar, al.AcceptanceDate,106) as [AcceptanceDate], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate] from " + A.AD.CustAccess + "_Allocation al inner join " + A.AD.CustAccess + "_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + A.AD.AutoArtID + "') and al.DeptCode=" + A.AD.DeptCode + " and al.Stage='" + A.AD.Stage + "'", strConnectionSite);
                //ds = DBProc.GetResultasDataSet("Select (select JBM_ID from JBM_INFO WHERE JBM_AutoID=ai.JBM_AutoID) as [JournalID],ai.ChapterID,ai.AutoArtID,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate],CONVERT(varchar, al.AcceptanceDate,106) as [AcceptanceDate],al.AcceptanceYN,ai.JBM_AutoID,al.AllocatedTo,(CASE WHEN al.AcceptanceDate is null THEN 'Yet to confirm' end) as [Remarks] from " + A.AD.CustAccess + "_Allocation al inner join " + A.AD.CustAccess + "_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + A.AD.AutoArtID + "') and al.DeptCode=" + A.AD.DeptCode + " and al.Stage='" + A.AD.Stage + "'", strConnectionSite);
                for (int i = 0; i <= strAutoArtIDs.Length-1; i++)
                {
                    string[] strSplitIDStage = strAutoArtIDs[i].Split('-');
                    //    ds = DBProc.GetResultasDataTbl(";WITH cte as (Select '' as RN, (select JBM_ID from JBM_INFO WHERE JBM_AutoID=ai.JBM_AutoID) as [JournalID],ai.ChapterID,ai.AutoArtID,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate],CONVERT(varchar, al.AcceptanceDate,106) as [AcceptanceDate],al.AcceptanceYN,ai.JBM_AutoID,al.AllocatedTo, (CASE WHEN al.AcceptanceDate is null THEN 'Yet to confirm' end) as [Remarks],null as [AccTime],al.Stage as [Stage] from TF_Allocation al inner join TF_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + strSplitIDStage[1].ToString().Replace("'","") + "') and al.DeptCode='" + A.AD.DeptCode + "' and al.Stage='" + strSplitIDStage[0].ToString().Replace("'", "") + "' and al.AllocatedTo='" + A.AD.UsrLogin + "' UNION ALL Select ROW_NUMBER()OVER(Partition by p.AutoArtId Order by  p.AccTime desc) As RN, (select JBM_ID from JBM_INFO WHERE JBM_AutoID = p.JBM_AutoID) as [JournalID],ai.ChapterID,ai.AutoArtID,'' as [AllocatedToName],'' as [FLDueDate],'' as [AcceptanceDate],'No' as [AcceptanceYN],ai.JBM_AutoID,p.EmpAutoID as [AllocatedTo], Description as [Remarks],AccTime,p.Stage as [Stage] from TF_ProcessInfo p inner join TF_ArticleInfo ai ON ai.AutoArtID = p.AutoArtID   where p.EmpAutoID = '" + A.AD.UsrLogin + "' and p.SubProcess = 'CE' and p.AutoArtID in ('" + strSplitIDStage[1].ToString().Replace("'", "") + "')  and Stage='" + strSplitIDStage[0].ToString().Replace("'", "") + "' and ProcessID='PID001') Select* FROM cte WHERE RN in (0, 1)", strConnectionSite);

                    ds = DBProc.GetResultasDataTbl(";WITH cte as (Select ROW_NUMBER()OVER(Partition by AutoArtId Order by  AccTime desc) As RN, * from  (Select (select JBM_ID from JBM_INFO WHERE JBM_AutoID=ai.JBM_AutoID) as [JournalID],ai.ChapterID,ai.AutoArtID,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate],FORMAT(al.AcceptanceDate, N'dd MMM yyyy hh:mm:ss:tt') as [AcceptanceDate],al.AcceptanceYN,ai.JBM_AutoID,al.AllocatedTo, (CASE WHEN al.AcceptanceDate is null THEN 'Yet to confirm' end) as [Remarks],al.AcceptanceDate as [AccTime],al.Stage as [Stage] from TF_Allocation al inner join TF_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + strSplitIDStage[1].ToString().Replace("'","") + "') and al.DeptCode='" + A.AD.DeptCode + "' and al.Stage='" + strSplitIDStage[0].ToString().Replace("'", "") + "' and al.AllocatedTo='" + A.AD.UsrLogin + "' UNION ALL Select (select JBM_ID from JBM_INFO WHERE JBM_AutoID = p.JBM_AutoID) as [JournalID],ai.ChapterID,ai.AutoArtID,'' as [AllocatedToName],'' as [FLDueDate],'' as [AcceptanceDate],'No' as [AcceptanceYN],ai.JBM_AutoID,p.EmpAutoID as [AllocatedTo], Description as [Remarks],AccTime,p.Stage as [Stage] from TF_ProcessInfo p inner join TF_ArticleInfo ai ON ai.AutoArtID = p.AutoArtID   where p.EmpAutoID = '" + A.AD.UsrLogin + "' and p.SubProcess = 'CE' and p.AutoArtID in ('" + strSplitIDStage[1].ToString().Replace("'", "") + "')  and Stage='" + strSplitIDStage[0].ToString().Replace("'", "") + "' and ProcessID='PID001') temp) Select * FROM cte WHERE RN in (1)", strConnectionSite);

                    if (ds.Rows.Count > 0)
                    {
                        dsFinal.Merge(ds);
                    }
                }
                


                var JSONString = from a in dsFinal.AsEnumerable()
                                 select new[] { CreateWebControl(a, "check"), a[1].ToString(), a[2].ToString(), a[3].ToString(), a[4].ToString(), a[5].ToString(), a[6].ToString(), a[7].ToString(), a[10].ToString()
                 };
                return Json(new { dataComp = JSONString}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return Json(new { dataNote = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult AcceptJobs(string CheckedArticleID)
        {
            try
            {
                string strConnectionSite = "dbConnSmartBG-LIVE";
                //string strConnectionSite = "dbConnSmartCH-LIVE";
                List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedArticleID);
                if (chkIds.Count > 0)
                {
                    string strAutoArtIDs = "";
                    string strEmpAutoID = "";
                    string strStage = "";
                    for (int i = 0; i < chkIds.Count; i++)
                    {
                        strAutoArtIDs = "'" + chkIds[i].Split('|')[0].ToString().Trim() + "'";
                        strEmpAutoID = chkIds[i].Split('|')[2].ToString().Trim();
                        strStage = chkIds[i].Split('|')[4].ToString().Trim();
                        bool result = DBProc.UpdateRecord("Update TF_Allocation Set AcceptanceYN='Yes', AcceptanceDate=GETDATE() where AutoArtID in (" + strAutoArtIDs + ") and Stage='" + strStage  + "' and DeptCode='40' and AllocatedTo='" + strEmpAutoID + "' and AcceptanceDate is null", strConnectionSite);
                    }

                    //strAutoArtIDs = strAutoArtIDs.Remove(strAutoArtIDs.Length - 1, 1).ToString();

               

                }
                return Json(new { dataJson = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return Json(new { dataJson = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult RejectJobs(string CheckedArticleID)
        {
            
            try
            {
                string strConnectionSite = "dbConnSmartBG-LIVE";
                //string strConnectionSite = "dbConnSmartCH-LIVE";
                List<string> chkIds = JsonConvert.DeserializeObject<List<string>>(CheckedArticleID);
                if (chkIds.Count > 0)
                {

                    string strMsgArticles = string.Empty;
                    string strMsgAutoArtID = string.Empty;
                    string strMsgEmpName = string.Empty;
                    string strMsgEmpAutoID = string.Empty;

                    for (int i = 0; i < chkIds.Count; i++)
                    {
                        string strAutoArtID = "";
                        string strEmpAutoID = "";
                        string strJBM_AutoID = "";
                        string strEmpName = "";
                        string strStage = "";
                        string strJNL = "";
                        strMsgAutoArtID = chkIds[i].Split('|')[0].ToString().Trim();
                        strAutoArtID = chkIds[i].Split('|')[0].ToString().Trim();
                        strEmpAutoID = chkIds[i].Split('|')[2].ToString().Trim();
                        strMsgEmpAutoID = chkIds[i].Split('|')[2].ToString().Trim();
                        strJBM_AutoID = chkIds[i].Split('|')[3].ToString().Trim();
                        strEmpName = chkIds[i].Split('|')[1].ToString().Trim();
                        strMsgEmpName = chkIds[i].Split('|')[1].ToString().Trim();
                        strStage = chkIds[i].Split('|')[5].ToString().Trim();
                        strJNL = chkIds[i].Split('|')[6].ToString().Trim();

                        strMsgArticles += "<tr><td>" + strJNL + "</td><td>" + chkIds[i].Split('|')[4].ToString().Trim() + "</td></tr>";

                        string strResult = DBProc.GetResultasString("INSERT INTO [TF_ProcessInfo] (EmpAutoID,AccTime,ProcessID,JBM_AutoID,AutoArtID,SubProcess,Stage,Description,ShortDescript) VALUES ('" + strEmpAutoID + "',GETDATE(),'PID001','" + strJBM_AutoID + "','" + strAutoArtID + "','CE','" + strStage + "','Job Rejected by " + strEmpName  + "','AcceptRejectScreen')", strConnectionSite);
                        strResult = DBProc.GetResultasString("DELETE FROM TF_Allocation WHERE AutoArtID='" + strAutoArtID + "' and DeptCode='40' and Stage='" + strStage + "' and AllocatedTo='" + strEmpAutoID + "'", strConnectionSite);

                        //If user rejected "Freelancer Rejected to CE",  "CE to Start"
                        bool result = DBProc.UpdateRecord("Update TF_StageInfo Set ArtStageTypeID='S1234_S1040_S1003_40385_" + DateTime.Now.ToString("dd-MM-yyyy h:mm tt") + "', CurrentStatus='CE to Start' where AutoArtID in ('" + strAutoArtID + "') and RevFinStage='" + strStage + "'", strConnectionSite);
                    }


                    string StrQry = "Select F.CustType,F.JBM_AutoID,F.Stage, MailFrom,e.EmpMailId as [MailTo],MailCc,MailBcc,M.Msgsubject,M.MsgBody from JBM_FTP_Stage_Details F inner join JBM_MessageInfo M ON M.MsgID=F.MailBody  join JBM_EmployeeMaster e on e.EmpAutoID='" + strMsgEmpAutoID + "' and F.CustType='TF' and F.Stage='CEReject'";
                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(StrQry, strConnectionSite);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string strJBMAutoID = ds.Tables[0].Rows[0]["JBM_AutoID"].ToString();
                        string strStage = ds.Tables[0].Rows[0]["Stage"].ToString();
                        string strMailfrom = ds.Tables[0].Rows[0]["MailFrom"].ToString();
                        string strMailTo = ds.Tables[0].Rows[0]["MailTo"].ToString();
                        string strMailCc = ds.Tables[0].Rows[0]["MailCc"].ToString();
                        string strMailBcc = ds.Tables[0].Rows[0]["MailBcc"].ToString();
                        string strMsgSubject = ds.Tables[0].Rows[0]["Msgsubject"].ToString();
                        string strMsgBody = ds.Tables[0].Rows[0]["MsgBody"].ToString();

                        strMsgBody = strMsgBody.Replace("###EmpName###", strMsgEmpName);
                        strMsgBody = strMsgBody.Replace("###ArticleIDS###", strMsgArticles);


                        string strMailTriggerQry = "INSERT INTO JBM_MailInfo ([EmpAutoID],[JBM_AutoID],[AutoArtID],[MailEventID],[MailStatus],[MailInitDate],[MailFrom],[MailTo],[MailCc],[MailBCc],[MailSub],[MailBody],[IsBodyHtml],[MailPriority],[MailType],[CustType],[Stage],[Mail_Attachment]) VALUES('E00001','" + strJBMAutoID + "','" + strMsgAutoArtID + "','CE0001','0',GETDATE(),'" + strMailfrom + "','" + strMailTo + "','" + strMailCc + "','','" + strMsgSubject + "','" + strMsgBody + "','1','2','Direct','TF','FP','')";
                        string strResult = DBProc.GetResultasString(strMailTriggerQry, strConnectionSite);
                    }
                    else {
                        return Json(new { dataJson = "Notification mail configuration is missing." }, JsonRequestBehavior.AllowGet);
                    }

                }
                return Json(new { dataJson = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return Json(new { dataJson = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult FLResponse(string info)
        {
            try
            {
                RL.ArtDet A = new RL.ArtDet();
                SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt

                //// Test
                string encryptString = objDS.EncryptData("BG|TF|E003498|40|FP-T299893,CER1-T304976,FP-T322057,FP-T336251,FP-T354247");

                info = HttpUtility.UrlDecode(info);
                info = info.Replace(" ", "+");
                string strreturnURL = objDS.DecryptData(info);

                //string[] strSplit = strreturnURL.Split('|');

                //A.AD.CustAccess = strSplit[0].ToString();
                //A.AD.UsrLogin = strSplit[1].ToString();
                //A.AD.JrnlSiteID = strSplit[2].ToString();
                //A.AD.Stage = strSplit[3].ToString();
                //A.AD.AutoArtID = strSplit[4].ToString().Replace(",","','");
                //A.AD.DeptCode = strSplit[5].ToString();
                //A.AD.DBLoc = strSplit[6].ToString();
                //A.AD.strTemp = strSplit[7].ToString();

                //string strConnectionSite = "dbConnSmart" + A.AD.DBLoc + "-LIVE";

                //DataSet ds = new DataSet();
                ////ds = DBProc.GetResultasDataSet("Select ai.ChapterID,ai.JBM_AutoID,ai.AutoArtID,al.AllocatedTo,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName],al.AcceptanceYN,CONVERT(varchar, al.AcceptanceDate,106) as [AcceptanceDate], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate] from " + A.AD.CustAccess + "_Allocation al inner join " + A.AD.CustAccess + "_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + A.AD.AutoArtID + "') and al.DeptCode=" + A.AD.DeptCode + " and al.Stage='" + A.AD.Stage + "'", strConnectionSite);
                //ds = DBProc.GetResultasDataSet("Select ai.JBM_AutoID,ai.AutoArtID,al.AllocatedTo,(Select EmpName from JBM_EmployeeMaster  WHERE EmpAutoID=al.AllocatedTo) as [AllocatedToName],al.AcceptanceYN,CONVERT(varchar, al.AcceptanceDate,106) as [AcceptanceDate], CONVERT(varchar, al.AllocatedDate,106) as [FLDueDate] from " + A.AD.CustAccess + "_Allocation al inner join " + A.AD.CustAccess + "_ArticleInfo ai ON ai.AutoArtID=al.AutoArtID  WHERE al.AutoArtID in ('" + A.AD.AutoArtID + "') and al.DeptCode=" + A.AD.DeptCode + " and al.Stage='" + A.AD.Stage + "'", strConnectionSite);

                //DataSet ds1 = new DataSet();
                //ds1 = DBProc.GetResultasDataSet("Select ChapterID, JBM_AutoID from " + A.AD.CustAccess + "_ArticleInfo WHERE AutoArtID='" + A.AD.AutoArtID + "'", strConnectionSite);

                //if (ds1.Tables[0].Rows.Count > 0)
                //{
                //    A.AD.ArticleID = ds1.Tables[0].Rows[0]["ChapterID"].ToString();
                //}


                //if (A.AD.strTemp == "Yes")
                //{
                //    if (ds.Tables[0].Rows.Count > 0)
                //    {
                //        if (ds.Tables[0].Rows[0]["AcceptanceYN"].ToString().ToLower().Trim() == "yes")
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " accepted on " + ds.Tables[0].Rows[0]["AcceptanceDate"].ToString() + ".";
                //        }
                //        else if (ds.Tables[0].Rows[0]["AcceptanceYN"].ToString().ToLower().Trim() == "no")
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " rejected on " + ds.Tables[0].Rows[0]["AcceptanceDate"].ToString() + ".";
                //        }
                //        else
                //        {
                //            bool result = DBProc.UpdateRecord("Update " + A.AD.CustAccess + "_Allocation Set AcceptanceDate=GETDATE(),AcceptanceYN='Yes' where AutoArtID='" + A.AD.AutoArtID + "' and Stage='" + A.AD.Stage + "' and DeptCode='" + A.AD.DeptCode + "'", strConnectionSite);

                //            if (result == true)
                //                Session["sAcceptReject"] = "Thanks for your confirmation. Article (" + A.AD.ArticleID  + ") has been assigned to you. The expected date of the edited document is (" + ds.Tables[0].Rows[0]["FLDueDate"].ToString() + ").";
                //            else
                //                Session["sAcceptReject"] = "Failed";

                //        }
                //    }
                //    else
                //    {
                //        ds = new DataSet();
                //        ds = DBProc.GetResultasDataSet("SELECT AutoArtID,CONVERT(varchar,AccTime,106) as [AccTime] from " + A.AD.CustAccess + "_ProcessInfo WHERE Processid='PID001' and AutoArtID='" + A.AD.AutoArtID + "' and SubProcess='CE' and Stage='FP'", strConnectionSite);

                //        if (ds.Tables[0].Rows.Count > 0)
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " rejected on " + ds.Tables[0].Rows[0]["AccTime"].ToString() + ". Please contact SPM for re-allocation.";
                //        }
                //        else
                //        {
                //            Session["sAcceptReject"] = "Allocation details not available for " + A.AD.ArticleID + ", Please contact SPM.";
                //        }
                //    }
                //}
                //else
                //{
                //    if (ds.Tables[0].Rows.Count > 0)
                //    {
                //        if (ds.Tables[0].Rows[0]["AcceptanceYN"].ToString().ToLower().Trim() == "yes")
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " accepted on " + ds.Tables[0].Rows[0]["AcceptanceDate"].ToString() + ".";
                //        }
                //        else if (ds.Tables[0].Rows[0]["AcceptanceYN"].ToString().ToLower().Trim() == "no")
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " rejected on " + ds.Tables[0].Rows[0]["AcceptanceDate"].ToString() + ".";
                //        }
                //        else
                //        {
                //            string strResult = DBProc.GetResultasString("DELETE FROM " + A.AD.CustAccess + "_Allocation WHERE  AllocatedTo='" + A.AD.UsrLogin + "' and AutoArtID='" + A.AD.AutoArtID + "' and Stage='" + A.AD.Stage + "' and DeptCode=" + A.AD.DeptCode + "", strConnectionSite);

                //            if (strResult != "")
                //            {
                //                Session["sAcceptReject"] = "Article (" + A.AD.ArticleID + ") has been rejected by you. This article will be pushed back in the allocation pool.";
                //                strResult = DBProc.InsertRecord("INSERT INTO " + A.AD.CustAccess + "_ProcessInfo (EmpAutoID,AccTime,ProcessID,JBM_AutoID,AutoArtID,SubProcess,Stage,Description) VALUES  ('" + A.AD.UsrLogin + "',GETDATE(),'PID001','" + ds.Tables[0].Rows[0]["JBM_AutoID"].ToString() + "','" + A.AD.AutoArtID + "','CE','" + A.AD.Stage + "','Job Rejected by " + ds.Tables[0].Rows[0]["AllocatedToName"].ToString() + "')", strConnectionSite);
                //            }
                //            else
                //            {
                //                Session["sAcceptReject"] = "Failed";
                //            }

                //        }
                //    }
                //    else
                //    {
                //        ds = new DataSet();
                //        ds = DBProc.GetResultasDataSet("SELECT AutoArtID,CONVERT(varchar,AccTime,106) as [AccTime] from " + A.AD.CustAccess + "_ProcessInfo WHERE Processid='PID001' and AutoArtID='" + A.AD.AutoArtID + "' and SubProcess='CE' and Stage='FP'", strConnectionSite);

                //        if (ds.Tables[0].Rows.Count > 0)
                //        {
                //            Session["sAcceptReject"] = "Article " + A.AD.ArticleID + " rejected on " + ds.Tables[0].Rows[0]["AccTime"].ToString() + ". Please contact SPM for re-allocation.";
                //        }
                //        else
                //        {
                //            Session["sAcceptReject"] = "Allocation details not available for " + A.AD.ArticleID + ", Please contact SPM.";
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                Session["sAcceptReject"] = "Failed: " +  ex.Message.ToString();
            }
            
            return View();
        }
        public string CreateWebControl(DataRow row, string strType)
        {

            string formControl = string.Empty;
            try
            {
                string uniqueID = row[3].ToString();
                string strJBMAutoID = row[8].ToString();
                string strJNL = row[1].ToString();
                string strEmpAutoID = row[9].ToString();
                string strStage = row[12].ToString();

                if (strType == "check")
                {
                    string isDisabled = "";
                    if (row[7].ToString() == "Yes" || row[7].ToString() == "No")
                    {
                        isDisabled= "disabled";
                    }

                    formControl = "<input type='checkbox' style='width: 18.8281px;' onClick=\"chkAllocationJob('" + uniqueID + "')\" class='caseChk' id='chk" + uniqueID + "' name='" + uniqueID + "' value='KGL' data-jnl='" + strJNL + "'  data-at='" + strJBMAutoID + "'  data-emp='" + strEmpAutoID + "' data-stage='" + strStage + "' " + isDisabled  + "/>";
                }
                //else if (strType == "Pickup")
                //{
                //    formControl = "<a href='javascript:void(0);'  onClick=\"btnPickupJob('" + uniqueID + "')\" class='btn btn-round btn-outline-secondary btn-sm' id='pickup" + uniqueID + "'  data-at='" + strJBMAutoID + "'>Pickup</a>";
                //}

                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }

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
                    articleIDList = ("'" + articleID + "'").Replace(",", "','");
                }

                if (articleIDList != "")
                {
                    articleIDList = " and (a.AutoArtID in (" + articleIDList + ") or a.ChapterID in (" + articleIDList + ") or a.IntrnlID in (" + articleIDList + "))";
                }

                string strQuery = "select '' as tblCheckBox,i.JBM_id as [Journal ID],a.ChapterID as [Job ID],a.AutoArtID,RevFinStage as Stage,dbo.fn_Currentstage(s.ArtstageTypeID) as [Current Process],s.Rev_Wf as [Work Flow],s.ArtstageTypeID from " + Session["sCustAcc"].ToString() + Init_Tables.gTblChapterOrArticleInfo + " a inner join " + Init_Tables.gTblJrnlInfo + " i on i.JBM_AutoID=a.JBM_AutoID inner join JBM_CustomerMaster c on c.CustID=i.CustID inner join " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " s on s.AutoArtID=a.AutoArtID where c.CustID = ISNULL(NULLIF(convert(varchar(10),'" + custID + "'),''),c.CustID) and i.JBM_AutoID = ISNULL(NULLIF(convert(varchar(10),'" + JBM_AutoID + "'),''),i.JBM_AutoID)  and s.DispatchDate is null " + articleIDList;

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

                        strQuery = strQuery + " Update " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " set Rev_Wf='" + WF + "' " + updateFields + ",CeArtStageTypeID=case when ISNULL(CeArtStageTypeID,'')<>'' then " + ArtStageTypeID + " else CeArtStageTypeID end, XMLArtStageTypeID = case when ISNULL(XMLArtStageTypeID,'') <> '' then " + ArtStageTypeID + " else XMLArtStageTypeID end " + " where autoartid='" + autoArtID + "' and RevFinStage='" + stage + "' and DispatchDate is null ";

                        strQuery = strQuery + " update " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " set ArtStageTypeID=" + ArtStageTypeID + " where AutoArtID='" + autoArtID + "' and Stage='" + stage + "' ";

                        if (currentStageR.Contains("S1017") | Regex.IsMatch(currentArtStageTypeID, "(S1026|S1050|S1051|S1052)"))
                        {
                            strQuery = strQuery + " delete from " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " where autoartid='" + autoArtID + "' and stage='" + stage + "' and deptCode not in('20','90','100') ";
                        }

                        strQuery = strQuery + " insert into " + Init_Tables.gTblProdAccess + "(AutoArtID,EmpAutoID,CustAcc,Descript,AccTime,AccPage,Process) values('" + autoArtID + "', '" + Session["EmpAutoID"].ToString() + "', '" + Session["sCustAcc"].ToString() + "', '" + trackingDesc + "', getdate(), 'Activity/WF Change', 'Activity/WF Change')";
                    }

                    //updateFields = updateFields + ",CeArtStageTypeID=case when ISNULL(CeArtStageTypeID,'')<>'' then " + ArtStageTypeID + " else CeArtStageTypeID end, XMLArtStageTypeID = case when ISNULL(XMLArtStageTypeID,'') <> '' then " + ArtStageTypeID + " else XMLArtStageTypeID end ";
                }

                //if(trackingQry != "")
                //{
                //    trackingQry = "insert into " + Init_Tables.gTblProdAccess + " (AutoArtID,EmpAutoID,CustAcc,Descript,AccTime,AccPage,Process) values" + trackingQry + "";
                //} 

                //string strQuery = "Update " + Session["sCustAcc"].ToString() + Init_Tables.gTblStageInfo + " set Rev_Wf='" + WF + "' " + updateFields + " where " + autoArtIDlist + "  and DispatchDate is null " + trackingQry + deleteAllocationQry + "";

                int result = DBProc.execNonQuery(strQuery, Session["sConnSiteDB"].ToString());
                //int result = 0;

                if (result == 1)
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