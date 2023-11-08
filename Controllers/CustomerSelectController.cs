using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using System.Text;
using System.Data;
using SmartTrack.Helper;
using System.Xml;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace SmartTrack.Controllers
{
    public class CustomerSelectController : Controller
    {
        // GET: CustomerSelect
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        clsCollection clsCollec = new clsCollection();

        CommonController objComm = new CommonController();

        [SessionExpire]
        public ActionResult Index()
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

            Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
            Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
            Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();

            string strConnectionSite;
            if (GlobalVariables.strSiteCode != null)
            {
                strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
            }
            else { strConnectionSite = "DefaultSiteDB"; }

            Session["sConnSiteDB"] = strConnectionSite;

            string strRestult = "";
            string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + Session["UserID"].ToString() + "' or a.EmpMailid = '" + Session["UserID"].ToString() + "') and a.emppass='" + Session["EmpPass"].ToString() + "' and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

            if (ds.Tables[0].Rows.Count > 0)
            {
                // To initialize session
                strRestult = objComm.InitializeSession(ds);
            }


            List<SelectListItem> lstCustomer = new List<SelectListItem>();
            ds = new DataSet();

            StringBuilder strEmployeeQuery = new StringBuilder(), strAccountType = new StringBuilder();
            strAccountType.Append("Exec('");
            strEmployeeQuery.Append("Declare @Cust as Varchar(200) Select @Cust = (RIGHT(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2), LEN(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2)) - 2)) from dbo.JBM_EmployeeMaster where EmpLogin = '" + Session["UserID"].ToString() + "' ");
            strAccountType.Append("Select CustAccess,AccountType from dbo.JBM_AccountTypeDesc where CustAccess in (' + @Cust + ')");
            strAccountType.Append("')");

            string strQueryFinal = strEmployeeQuery.ToString() + strAccountType.ToString();

            ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                {
                    lstCustomer.Add(new SelectListItem
                    {
                        Value = ds.Tables[0].Rows[intCount]["CustAccess"].ToString(),
                        Text = ds.Tables[0].Rows[intCount]["AccountType"].ToString()
                    });
                }

            }

            if (Regex.IsMatch(Session["UserID"].ToString(), "(40385|45817|43637|44310)", RegexOptions.IgnoreCase))
            {
                lstCustomer.Add(new SelectListItem
                {
                    Value = "Add Customer",
                    Text = "Add Customer"
                });
            }


            ViewBag.Customerlist = lstCustomer;

            //ViewBag.BindCustomer = BindCustomers(Session["sSiteID"].ToString());
            return View();
        }
        [SessionExpire]
        public JsonResult LoadCustomer(string vSiteID)
        {
            try
            {
               
                string strConnectionSite = "";
                strConnectionSite = "dbConnSmart" + vSiteID + "-" + GlobalVariables.strEnvironment.Trim().ToString();

                Session["sConnSiteDB"] = strConnectionSite;
                if (vSiteID == "CH")
                {
                    GlobalVariables.strSiteID = "L0001";
                    GlobalVariables.strSiteCode = "CH";
                    GlobalVariables.strSiteLocation = "Chennai";
                }
                else if (vSiteID == "MB")
                {
                    GlobalVariables.strSiteID = "L0002";
                    GlobalVariables.strSiteCode = "MB";
                    GlobalVariables.strSiteLocation = "Mumbai";
                }
                else if (vSiteID == "ND")
                {
                    GlobalVariables.strSiteID = "L0003";
                    GlobalVariables.strSiteCode = "ND";
                    GlobalVariables.strSiteLocation = "Noida";
                }
                else if (vSiteID == "BG")
                {
                    GlobalVariables.strSiteID = "L0004";
                    GlobalVariables.strSiteCode = "BG";
                    GlobalVariables.strSiteLocation = "Bangalore";
                }

                Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
                Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
                Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();

                //Session reassign

                string strRestult = "";
                string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + Session["UserID"].ToString()  + "' or a.EmpMailid = '" + Session["UserID"].ToString() + "') and a.emppass='" + Session["EmpPass"].ToString() + "' and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {

                    strRestult = objComm.InitializeSession(ds);

                    ds = new DataSet();

                    StringBuilder strEmployeeQuery = new StringBuilder(), strAccountType = new StringBuilder();
                    strAccountType.Append("Exec('");
                    strEmployeeQuery.Append("Declare @Cust as Varchar(200) Select @Cust = (RIGHT(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2), LEN(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2)) - 2)) from dbo.JBM_EmployeeMaster where EmpLogin = '" + Session["UserID"].ToString() + "' ");
                    strAccountType.Append("Select CustAccess,AccountType from dbo.JBM_AccountTypeDesc where CustAccess in (' + @Cust + ')");
                    strAccountType.Append("')");

                    string strQueryFinal = strEmployeeQuery.ToString() + strAccountType.ToString();

                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    if (Regex.IsMatch(Session["UserID"].ToString(), "(40385|45817|43637|44310)", RegexOptions.IgnoreCase))
                    {
                        DataRow workRow;
                        workRow = ds.Tables[0].NewRow();
                        workRow[0] = "Add Customer";
                        workRow[1] = "Add Customer";
                        ds.Tables[0].Rows.Add(workRow);
                    }

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        var JSONString = from a in ds.Tables[0].AsEnumerable()
                                         select new[] {a[0].ToString(), a[1].ToString()
                                         };

                        return Json(new { aaData = JSONString }, JsonRequestBehavior.AllowGet);

                    }
                    else {
                        return Json(new { aaData = "NoCustAccess" }, JsonRequestBehavior.AllowGet);
                    }
                }
                else{
                    Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
                    Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
                    Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();
                    return Json(new { aaData = "NoSiteAccess" }, JsonRequestBehavior.AllowGet);

                }

              //  return View();
                // return RedirectToAction("Index", "CustomerSelect");
            }
            catch (Exception ex)
            {
                return Json(new { aaData = "Failed", abData = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        private string BindCustomers(string siteID)
        {
            StringBuilder strEmployeeQuery = new StringBuilder(), strAccountType = new StringBuilder();
            strAccountType.Append("Exec('");
            strEmployeeQuery.Append("Declare @Cust as Varchar(200) Select @Cust = (RIGHT(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2), LEN(LEFT(Replace(CustAccess, '|', ''','''), LEN(Replace(CustAccess, '|', ''','''))-2)) - 2)) from dbo.JBM_EmployeeMaster where EmpLogin = '" + Session["UserID"].ToString() + "' ");
            strAccountType.Append("Select CustAccess,AccountType from dbo.JBM_AccountTypeDesc where CustAccess in (' + @Cust + ')");
            strAccountType.Append("')");

            string strQuery = strEmployeeQuery.ToString() + strAccountType.ToString();

            DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

            if (dt.Rows.Count != 0)
            {
                string strTable = "<div id='tblCust' style='width:100%;margin-top: 150px;'><table id='tblCustomers' cellpadding='0' cellspacing='0' runat='server' style='width: 100%' runat='server'>";
                string StrOut = "";
                string strRowOpen = "<tr style='width: 99%'>";
                string strRowClose = "</tr>";
                int i = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string strCustAccess = dr["CustAccess"].ToString();
                    string strAccountName = dr["AccountType"].ToString();

                    // 'Dim strLink As String = "<td style='width: 25%; padding-bottom:10px' align='center'><div class='custDiv'><a href='#' onclick='goToCustomer(""" + strCustAccess + """, """ + StrLoc + """)' class='btn btn-instagram' style='color: #0A94D0; border-radius: 3px; text-align: center; background-position: center; background-color: rgba(255, 255, 255, 0.2); transition: background-color .1s .2s; width: 90%;display: inline;padding: 4px 12px;vertical-align: middle;background-image: none; text-shadow: none; font: normal 15px auto Verdana, Arial;'><i class='material-icons'>&#xe7ef;</i>" & strAccountName & "</a></div></td>"

                    // '//By Sivakumar M on Oct 3rd 2019
                    int tblPercent = 20;
                    string divCls = "custDiv";
                    if (dt.Rows.Count <= 3)
                        tblPercent = 99 / dt.Rows.Count + 1 ;
                    else
                        tblPercent = 20;

                    if (dt.Rows.Count == 1)
                        divCls = "custDiv1";
                    else if (dt.Rows.Count == 2 | dt.Rows.Count == 3)
                        divCls = "custDiv";

                    

                    string strLink = "<td style='width: " + tblPercent + "%; padding-bottom:20px' align='center'><div class='" + divCls + "'><img title='" + strAccountName.Replace("&","") + "' style='cursor: pointer;' onclick='goToCustomer(\"" + strCustAccess + "\", \"" + siteID + "\")' src='../Images/Customer/" + strAccountName.Replace("&", "") + ".png'/></div><span hidden>" + strAccountName.Replace("&", "") + "</span></td>";

                    if (i == 1)
                        StrOut += strRowOpen + strLink;
                    else if (i == 5)
                    {
                        StrOut += strLink + strRowClose;
                        i = 0;
                    }
                    else
                        StrOut += strLink;
                    i += 1;
                }
                if (i != 1)
                    StrOut += "<td></td>" + strRowClose;

                return strTable + StrOut + " </table></div>";
            }
            else
            {
                return "";
            }

            
        }
        [SessionExpire]
        public ActionResult redirectToCustomerPage(string CustAcc)
        {
            string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + Session["UserID"].ToString() + "' or a.EmpMailid = '" + Session["UserID"].ToString() + "') and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
            DataSet ds = new DataSet();

            ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

            if (ds.Tables[0].Rows.Count > 0)
            {
                Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();
                Session["gTeamID"] = ds.Tables[0].Rows[0]["TeamID"].ToString();

                if (Session["gTeamID"].ToString() != "")
                {
                    clsInit.Proc_SplitTeam();
                }

                if (CustAcc != "")
                {
                    string strAccType = "Select ATD.AccountType, ATD.CustGroup, ATD.JBM_TeamID, ATD.Job_ID, (Select (case when ATD.RootID IS NULL then '' else R.RootPath end) from " + Init_Tables.gTblRootDirectory + " R where R.RootID=ATD.RootID) as RootDirPath, (Select (case when ATD.InwardDir IS NULL then '' else R.RootPath end) from " + Init_Tables.gTblRootDirectory + " R where R.RootID=ATD.InwardDir) as InwardDirPath, ATD.CustID, ATD.MLDir, ATD.GraphicsDirPath, ATD.MyPetInDir, ATD.MyPetOutDir, ATD.InwardFPDirPath, ATD.InwardRevDirPath, ATD.InwardIssueDirPath, ATD.DisplayCustIDPD, ATD.CreateDirectory, ATD.ImFpDirPath, ATD.ImRevDirPath, ATD.ImIssueDirPath, ATD.CeFpDirPath, ATD.CeRevDirPath, ATD.CeIssueDirPath, ATD.ReFpDirPath, ATD.ReRevDirPath, ATD.ReIssueDirPath, ATD.MLFpDirPath, ATD.MLRevDirPath, ATD.MLIssueDirPath, ATD.PagFpDirPath, ATD.PagRevDirPath, ATD.PagIssueDirPath, ATD.PrFpDirPath, ATD.PrRevDirPath, ATD.PrIssueDirPath, ATD.DispatchFpDirPath, ATD.DispatchRevDirPath, ATD.DispatchIssueDirPath, ATD.EproofCwpDir, ATD.DispatchVolDirPath, ATD.ErrorDirPath, ATD.CommentEnableDir, ATD.EproofNormal, ATD.DispatchUploadIn, ATD.DispatchSproofDirPath, ATD.SiteId, ATD.GraphicsOnlineDirPath, ATD.GraphicsWebDirPath, ATD.MarkPDFDirPath, ATD.CWP_CreatorEmail, ATD.IssueManagement, ATD.WorkingFolderDirPath, ATD.IceDirPath, ATD.DispatchHTML, ATD.ConvFpDirPath  from " + Init_Tables.gTblAccountTypeDesc + " ATD where ATD.CustAccess='" + CustAcc + "'";
                    DataSet dsAccType = DBProc.GetResultasDataSet(strAccType, Session["sConnSiteDB"].ToString()); //dbConnSmartTrack

                    if (dsAccType.Tables[0].Rows.Count > 0)
                    {
                        Session["sCustGroup"] = dsAccType.Tables[0].Rows[0]["CustGroup"];
                        Session["sCustTeamID"] = dsAccType.Tables[0].Rows[0]["JBM_TeamID"];
                        Session["AccountType"] = dsAccType.Tables[0].Rows[0]["AccountType"].ToString();
                    }
                }

                Session["sCustAcc"] = CustAcc;
                Session["UserID"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
            }


            string strConnectionSite = "";
            if (CustAcc == "TF")
            {
                strConnectionSite = "dbConnSmartBG" + "-" + GlobalVariables.strEnvironment.Trim().ToString();
                Session["sConnSiteDB"] = strConnectionSite;
                GlobalVariables.strSiteID = "L0004";
                GlobalVariables.strSiteCode = "BG";
                GlobalVariables.strSiteLocation = "Bangalore";

                Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
                Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
                Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();
            }


            var routeValues = new RouteValueDictionary { { "CustAcc", CustAcc }, { "EmpID", Session["UserID"].ToString() }, { "SiteID", GlobalVariables.strSiteID.ToString() } };
            if (CustAcc == "SG")
            {
                return RedirectToAction("SageSchedule", "Schedule", routeValues);
            }
            else {
                return RedirectToAction("Index", "StaffInbox", routeValues);
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