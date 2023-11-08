using System;
using System.Text;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
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
using System.Web.Routing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RL = ReferenceLibrary;

namespace SmartTrack.Controllers
{
    [Authorize]
    public class AccountController : Controller
    {
        private ApplicationSignInManager _signInManager;
        private ApplicationUserManager _userManager;

        DataProc DBProc = new DataProc(); // Data store/retrive DB
        Generic gen = new Generic();    
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt

        public AccountController()
        {

            //string strExcelPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Support\\SAGE-SMART Country.csv");
            //DataTable dt = new DataTable();
            //dt = DBProc.ReadCsvFile(strExcelPath);
            
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
        }

        public AccountController(ApplicationUserManager userManager, ApplicationSignInManager signInManager )
        {
            UserManager = userManager;
            SignInManager = signInManager;
        }

        public ApplicationSignInManager SignInManager
        {
            get
            {
                return _signInManager ?? HttpContext.GetOwinContext().Get<ApplicationSignInManager>();
            }
            private set 
            { 
                _signInManager = value; 
            }
        }

        public ApplicationUserManager UserManager
        {
            get
            {
                return _userManager ?? HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            }
            private set
            {
                _userManager = value;
            }
        }

        //
        // GET: /Account/Login
        [AllowAnonymous]
        public ActionResult Login(string returnUrl)
        {
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        //
        // POST: /Account/Login
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Login(LoginViewModel model, string returnUrl)
        {
            if (ModelState.IsValid)
            {
                //string strPassword = objDS.Encrypt(model.Password.ToString(), "*!%$@~&#?,:"); //&%#@?,:*
                string strPassword = ASCIItoHex(model.Password.ToString()); //&%#@?,:*

                string strConnectionSite = "";
                string strRequestURL = Request.Url.AbsoluteUri.ToString().ToLower();

                if (GlobalVariables.strSiteCode != null)
                {
                    strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() +  "-" +  GlobalVariables.strEnvironment.Trim().ToString();
                }
                else { strConnectionSite = "DefaultSiteDB"; }

                Session["sConnSiteDB"] = strConnectionSite;

                string strRestult = "";
                string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno,a.NASAccessCmd from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + model.UserID + "' or a.EmpMailid = '" + model.UserID + "') and a.emppass='" + strPassword + "' and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                //Write log
                gen.WriteLog("Login " + ds.Tables[0].Rows.Count);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    // To initialize session
                    strRestult = InitializeSession(ds);

                    DBProc.UpdateRecord("Update JBM_EmployeeMaster Set LastLogin = GETDATE() where EmpLogin = '" + Session["UserID"].ToString() + "'", Session["sConnSiteDB"].ToString());   //, Session_ID = '" + Session.SessionID + "'

                    if (Regex.IsMatch(ds.Tables[0].Rows[0]["EmpLogin"].ToString().ToLower(), "(pnas|ada|aapg|asn|aai)"))
                    {
                        Session["newSmartTrack"] = "pnas";
                        return RedirectToAction("ArticleReport", "Reports");
                    }

                    if (Regex.IsMatch(strRequestURL.ToLower(), "(smarttrack\\.kwglobal\\.com\\/smarttrack\\-kgl|smarttrack\\.kwglobal\\.com\\/smarttrack\\-uat|localhost)"))
                    {
                        return RedirectToAction("Index", "CustomerSelect");
                    }
                    else if (Regex.IsMatch(strRequestURL.ToLower(), "(smarttrack\\.kwglobal\\.com\\/smart\\-track)"))
                    {
                        if (Regex.IsMatch(ds.Tables[0].Rows[0]["RoleID"].ToString(), "(102)") && Regex.IsMatch(ds.Tables[0].Rows[0]["DeptCode"].ToString(), "(100|180|40|260)"))
                        {
                            //var routeValues = new RouteValueDictionary { { "CustAcc", "TF" }, { "sCustSN", "TandF" }, { "EmpID", model.UserID }, { "SiteID", "L0001" } };
                            Session["sCustAcc"] = "KW";
                            return RedirectToAction("RupDashboard", "Reports");
                            //return RedirectToAction("CompReport", "Reports", routeValues);
                        }
                        else
                        {
                            return RedirectToAction("Index", "CustomerSelect");
                        }

                    }
                    else if (Regex.IsMatch(strRequestURL.ToLower(), "(smarttrack\\.kwglobal\\.com\\/smarttrack\\-nd|smarttrack\\.kwglobal\\.com\\/smarttrack\\-nd\\-test|localhost)"))   /// To load Cenpro pages
                    {
                        if (Regex.IsMatch(ds.Tables[0].Rows[0]["DeptCode"].ToString(), "(10|40|140|260)"))
                        {
                            if (ds.Tables[0].Rows[0]["RoleID"].ToString() == "102")
                            {
                                Session["CustomerSN"] = model.UserID.ToUpper();
                                Session["CustName"] = "";
                                var routeValues = new RouteValueDictionary { { "CustAcc", "BK" }, { "sCustSN", model.UserID.ToUpper() }, { "EmpID", model.UserID }, { "SiteID", "L0003" } };
                                //return RedirectToAction("Index", "ProjectMgnt", routeValues);
                                return RedirectToAction("Dashboard", "ProjectTrack", routeValues);
                            }
                            else
                            {
                                Session["CustomerSN"] = "";
                                Session["CustName"] = "";
                                var routeValues = new RouteValueDictionary { { "CustAcc", "BK" }, { "EmpID", model.UserID }, { "SiteID", "L0003" } };
                                //return RedirectToAction("Index", "ProjectMgnt", routeValues);
                                return RedirectToAction("ProjectTracking", "ProjectTrack", routeValues);
                            }
                        }
                        else
                        {
                            return RedirectToAction("Dashboard", "Home");
                            //return RedirectToAction("CustomerSelect", "Index");
                        }
                    }
                    else
                    {
                        return RedirectToAction("Index", "CustomerSelect");
                    }


                }
                else {
                    ViewBag.StatusLogin = "InvalidPwd";
                    return View(model);
                }

                //if (model.Password == "SuperAdmin")
                //{
                //    System.Web.HttpContext.Current.Session["UserID"] = model.UserID;
                //    return RedirectToAction("Dashboard", "Home");
                //}
                //else if (model.Password == "Admin")
                //{
                //    return RedirectToAction("Contact", "Home");
                //}
                //else if (model.Password == "User")
                //{
                //    return RedirectToAction("About", "Home");
                //}
                //else
                //{
                //    return RedirectToAction("ForgotPassword", "Account");
                //}
            }
            else
            { return View(model); }

            // This doesn't count login failures towards account lockout
            // To enable password failures to trigger account lockout, change to shouldLockout: true
            var result = await SignInManager.PasswordSignInAsync(model.UserID, model.Password, model.RememberMe, shouldLockout: false);
            switch (result)
            {
                case SignInStatus.Success:
                    return RedirectToLocal(returnUrl);
                case SignInStatus.LockedOut:
                    return View("Lockout");
                case SignInStatus.RequiresVerification:
                    return RedirectToAction("SendCode", new { ReturnUrl = returnUrl, RememberMe = model.RememberMe });
                case SignInStatus.Failure:
                default:
                    ModelState.AddModelError("", "Invalid login attempt.");
                    return View(model);
            }
        }
        public string InitializeSession(DataSet ds)
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
                    Session["EmpMailId"] = ds.Tables[0].Rows[0]["EmpMailId"].ToString();
                    Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                    Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                    Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();
                    Session["gJwAccItm"] = ds.Tables[0].Rows[0]["JwAccessItm"].ToString();
                    Session["gTeamID"] = ds.Tables[0].Rows[0]["TeamID"].ToString();
                    Session["DeptAcc"] = ds.Tables[0].Rows[0]["DeptAccess"].ToString();
                    Session["AccessRights"] = "";
                    Session["CustomerSN"] = "";
                    Session["CustName"] = "";
                    Session["strHomeURL"] = "";
                    Session["sSiteID"] = GlobalVariables.strSiteID.ToString();
                    Session["sSiteCode"] = GlobalVariables.strSiteCode.ToString();
                    Session["sSiteLocation"] = GlobalVariables.strSiteLocation.ToString();
                    Session["newSmartTrack"] = "Yes";
                    Session["NASAccessCmd"] = ds.Tables[0].Rows[0]["NASAccessCmd"].ToString();

                    string CustAccess = ds.Tables[0].Rows[0]["CustAccess"].ToString();
                    if (CustAccess != "")
                    {

                        CustAccess = CustAccess.Substring(1, CustAccess.IndexOf("|", 1) - 1);
                        Session["sCustAcc"] = CustAccess;
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
        public string HexToASCII(string hexString)
        {
            String ascii = "";

            for (int i = 0; i < hexString.Length; i += 2)
            {
                String part = hexString.Substring(i, 2);

                char ch = (char)Convert.ToInt32(part, 16); ;
                ascii = ascii + ch;
            }
            return ascii;
        }
        public string ASCIItoHex(string Value)
        {
            StringBuilder sb = new StringBuilder();

            foreach (byte b in Value)
            {
                sb.Append(string.Format("{0:x2}", b));
            }

            return sb.ToString();
        }

        //To get the profile image
        [HttpPost]
        [AllowAnonymous]
        public ActionResult getProfileImage(string userId)
        {
            string strConnectionSite = "";
            string strRequestURL = Request.Url.AbsoluteUri.ToString().ToLower();

            if (GlobalVariables.strSiteCode != null)
            {
                strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
            }
            else { strConnectionSite = "DefaultSiteDB"; }

            string strUrl = Request.Url.AbsoluteUri.ToString().Replace("/Account/getProfileImage", "");

            String userAgent;
            userAgent = Request.UserAgent;
            if (userAgent.ToLower().IndexOf("mac") > -1)
            {
                Session["OSName"] = "MAC";
            }
            else { Session["OSName"] = "WINDOWS"; }
             
            string strQuery = "Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + userId + "' or a.EmpMailid = '" + userId + "')  and a.EmpStatus = '1'"; // and  a.EmailVerify = '1'
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet(strQuery, strConnectionSite);

            string imguserId;

            if (ds.Tables[0].Rows.Count > 0)
            {
                imguserId = ds.Tables[0].Rows[0][1].ToString();

                string strPath = System.Web.Hosting.HostingEnvironment.MapPath(@"~/Images/Employee/" + imguserId + ".png");
                
                if (System.IO.File.Exists(strPath))
                {
                    return Json(new { aaData = strUrl + "/Images/Employee/" + imguserId + ".png" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    strPath = strPath.Replace(".png", ".jpg");
                    if (System.IO.File.Exists(strPath))
                    {
                        return Json(new { aaData = strUrl + "/Images/Employee/" + imguserId + ".jpg" }, JsonRequestBehavior.AllowGet);
                    }
                    else {
                        strPath = strPath.Replace(".jpg", ".gif");
                        if (System.IO.File.Exists(strPath))
                        {
                            return Json(new { aaData = strUrl + "/Images/Employee/" + imguserId + ".gif" }, JsonRequestBehavior.AllowGet);
                        }
                        return Json(new { aaData =  strUrl + "/Images/Employee/profile.png" }, JsonRequestBehavior.AllowGet);
                    }
                    
                }

                //if (Regex.IsMatch(ds.Tables[0].Rows[0][1].ToString(), "(" + userId + ")", RegexOptions.IgnoreCase))
                //{}

            }
            else
            {
                return Json(new { aaData = "NoRecord" }, JsonRequestBehavior.AllowGet);
            }

            return Json(new { aaData = strUrl + "/Images/Employee/" + imguserId + ".gif" }, JsonRequestBehavior.AllowGet);

        }

        //
        // GET: /Account/VerifyCode
        [AllowAnonymous]
        public async Task<ActionResult> VerifyCode(string provider, string returnUrl, bool rememberMe)
        {
            // Require that the user has already logged in via username/password or external login
            if (!await SignInManager.HasBeenVerifiedAsync())
            {
                return View("Error");
            }
            return View(new VerifyCodeViewModel { Provider = provider, ReturnUrl = returnUrl, RememberMe = rememberMe });
        }

        //
        // POST: /Account/VerifyCode
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> VerifyCode(VerifyCodeViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            // The following code protects for brute force attacks against the two factor codes. 
            // If a user enters incorrect codes for a specified amount of time then the user account 
            // will be locked out for a specified amount of time. 
            // You can configure the account lockout settings in IdentityConfig
            var result = await SignInManager.TwoFactorSignInAsync(model.Provider, model.Code, isPersistent:  model.RememberMe, rememberBrowser: model.RememberBrowser);
            switch (result)
            {
                case SignInStatus.Success:
                    return RedirectToLocal(model.ReturnUrl);
                case SignInStatus.LockedOut:
                    return View("Lockout");
                case SignInStatus.Failure:
                default:
                    ModelState.AddModelError("", "Invalid code.");
                    return View(model);
            }
        }

        //
        // GET: /Account/Register
        [AllowAnonymous]
        public ActionResult Register()
        {
            var objCollec = new clsCollection();
            var model = new RegisterViewModel();
            model.Departments = objCollec.GetDepartments(Session["sConnSiteDB"].ToString());
            return View(model);

        }

        //
        // POST: /Account/Register
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Register(RegisterViewModel model)
        {

            if (ModelState.IsValid)
            {

                string strConnectionSite = "";
                string strRequestURL = Request.Url.AbsoluteUri.ToString().ToLower();

                if (GlobalVariables.strSiteCode != null)
                {
                    strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
                }
                else { strConnectionSite = "DefaultSiteDB"; }

                ///**************Need to check existing and write procedure 
                string SelectedValue = model.Department;

                //string strQuery = "INSERT INTO JBM_Employeemaster (EmpAutoID,EmpLogin,EmpPass,EmpName,EmpSurname,DeptCode, EmpMailId) values ('E' + (SELECT FORMAT(Count(*) + 1, REPLICATE('0', 6)) from JBM_Employeemaster), '" + model.EmailUserIID + "','" + model.Password + "','" + model.firstName + "','" + model.lastName + "'," + Convert.ToInt32(SelectedValue) + ",'" + model.Email + "')";
                //string strStatus = DBProc.InsertRecord(strQuery, "dbConnSmartTrack");
                
                string strPassword = objDS.Encrypt(model.Password.ToString(), "*!%$@~&#?,:"); //&%#@?,:*



                DataSet ds = new DataSet();

                SqlConnection con = new SqlConnection();
                con = DBProc.getConnection(strConnectionSite);
                con.Open();
                SqlCommand cmd = new SqlCommand("EmployeeRegister", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@EmpLogin", model.EmailUserIID);
                cmd.Parameters.AddWithValue("@EmpPass", strPassword);
                cmd.Parameters.AddWithValue("@EmpName", model.firstName);
                cmd.Parameters.AddWithValue("@EmpSurname", model.lastName);
                cmd.Parameters.AddWithValue("@DeptCode", Convert.ToInt32(SelectedValue));
                cmd.Parameters.AddWithValue("@EmpMailId", model.Email);
                cmd.Parameters.AddWithValue("@RegistrType", "Application");
                cmd.Parameters.AddWithValue("@EVerify", "2");
                //cmd.ExecuteNonQuery();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(ds);

                con.Close();

                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0][0].ToString() == "Success")
                    {
                        string strMessage = "Hi "+ model.firstName + ",<br/><br/>Thanks for signing up for Smart Publisher. You are just one step away from activation your account on the Smart Publisher.<br/><br/>Please verify the link <a href='https://localhost:44358/Home/AccountVerification/" + ds.Tables[0].Rows[0][4].ToString() + "?verify=" + ds.Tables[0].Rows[0][7].ToString() + "'>Activate Account</a><br/><br/>Regards,<br/>Smart Publisher<br/><br/>This is an auto generated mail.";
                        gen.Mail_Send("sivakumar.m@cenveo.com", "sivakumar.m@cenveo.com", "", "Smart Publisher Account Registration", strMessage);
                        return RedirectToAction("AccountRegistered", "Home", new { @id = "1" });
                    }
                    else
                    {
                        ViewBag.strStatus = "Already Exists";

                        var objCollecs = new clsCollection();
                        var remodels = new RegisterViewModel();
                        remodels.Departments = objCollecs.GetDepartments(Session["sConnSiteDB"].ToString());
                        return View(remodels);
                        //return RedirectToAction("Login", "Account");
                    }
                }
                        
                //var user = new ApplicationUser { UserName = model.Email, Email = model.Email };
                //var result = await UserManager.CreateAsync(user, model.Password);
                //if (result.Succeeded)
                //{
                //    await SignInManager.SignInAsync(user, isPersistent: false, rememberBrowser: false);

                //    // For more information on how to enable account confirmation and password reset please visit https://go.microsoft.com/fwlink/?LinkID=320771
                //    // Send an email with this link
                //    // string code = await UserManager.GenerateEmailConfirmationTokenAsync(user.Id);
                //    // var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = user.Id, code = code }, protocol: Request.Url.Scheme);
                //    // await UserManager.SendEmailAsync(user.Id, "Confirm your account", "Please confirm your account by clicking <a href=\"" + callbackUrl + "\">here</a>");

                //    return RedirectToAction("Index", "Home");
                //}

                //AddErrors(result);
            }

            // If we got this far, something failed, redisplay form
            var objCollec = new clsCollection();
            var remodel = new RegisterViewModel();
            remodel.Departments = objCollec.GetDepartments(Session["sConnSiteDB"].ToString());
            return View(remodel);
        }

        
        //
        // GET: /Account/ConfirmEmail
        [AllowAnonymous]
        public async Task<ActionResult> ConfirmEmail(string userId, string code)
        {
            if (userId == null || code == null)
            {
                return View("Error");
            }
            var result = await UserManager.ConfirmEmailAsync(userId, code);
            return View(result.Succeeded ? "ConfirmEmail" : "Error");
        }

        //
        // GET: /Account/ForgotPassword
        [AllowAnonymous]
        public ActionResult ForgotPassword()
        {
            return View();
        }

        //
        // POST: /Account/ForgotPassword
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> ForgotPassword(ForgotPasswordViewModel model)
        {
            if (ModelState.IsValid)
            {

                string strConnectionSite = "";
                string strRequestURL = Request.Url.AbsoluteUri.ToString().ToLower();

                if (GlobalVariables.strSiteCode != null)
                {
                    strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
                }
                else { strConnectionSite = "DefaultSiteDB"; }

                string strQuery = "Select a.EmpLogin, a.EmpPass, a.EmpName, a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE (a.Emplogin = '" + model.UserID + "' and a.EmpMailid = '" + model.Email + "') and a.EmpStatus = '1'";  // and a.EmailVerify = '1'
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, strConnectionSite); //dbConnSmartTrack
                if (ds.Tables[0].Rows.Count > 0)
                {
                    string strPassword;
                    string strEmail = ds.Tables[0].Rows[0]["EmpMailId"].ToString();
                    if (strEmail == model.Email.ToString().Trim())
                    {
                        strPassword = HexToASCII(ds.Tables[0].Rows[0]["EmpPass"].ToString()); //&%#@?,:*
                        string strMessage = "Hi " + ds.Tables[0].Rows[0]["EmpName"].ToString() + ",<br/><br/>Please find your password below.<br/><br/>Password: " + strPassword + "<br/><br/>Regards,<br/>Smart Track<br/><br/>This is an auto generated mail.";
                        //gen.Mail_Send("MIS.Support@kwglobal.com", strEmail, "sivakumar.m@kwglobal.com", "Smart Track - Forgot Password", strMessage);
                        gen.Mail_Send("MIS.Support@kwglobal.com", strEmail, "", "Smart Track - Forgot Password", strMessage);
                        ViewBag.Status = "Please check your email to know your password.";
                    }
                    else
                    {
                        ViewBag.Status = "Enter valid email ID";
                    }
                  }
                else
                {
                    ViewBag.Status = "Enter valid email ID";
                }
                //var user = await UserManager.FindByNameAsync(model.Email);
                //if (user == null || !(await UserManager.IsEmailConfirmedAsync(user.Id)))
                //{
                //    // Don't reveal that the user does not exist or is not confirmed
                //    return View("ForgotPasswordConfirmation");
                //}

                //// For more information on how to enable account confirmation and password reset please visit https://go.microsoft.com/fwlink/?LinkID=320771
                //// Send an email with this link
                //// string code = await UserManager.GeneratePasswordResetTokenAsync(user.Id);
                //// var callbackUrl = Url.Action("ResetPassword", "Account", new { userId = user.Id, code = code }, protocol: Request.Url.Scheme);		
                //// await UserManager.SendEmailAsync(user.Id, "Reset Password", "Please reset your password by clicking <a href=\"" + callbackUrl + "\">here</a>");
                //// return RedirectToAction("ForgotPasswordConfirmation", "Account");
            }

            // If we got this far, something failed, redisplay form
            return View(model);
        }

        //
        // GET: /Account/ForgotPasswordConfirmation
        [AllowAnonymous]
        public ActionResult ForgotPasswordConfirmation()
        {
            return View();
        }

        //
        // GET: /Account/ResetPassword
        [AllowAnonymous]
        public ActionResult ResetPassword(string code)
        {
            return code == null ? View("Error") : View();
        }

        //
        // POST: /Account/ResetPassword
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> ResetPassword(ResetPasswordViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }
            var user = await UserManager.FindByNameAsync(model.Email);
            if (user == null)
            {
                // Don't reveal that the user does not exist
                return RedirectToAction("ResetPasswordConfirmation", "Account");
            }
            var result = await UserManager.ResetPasswordAsync(user.Id, model.Code, model.Password);
            if (result.Succeeded)
            {
                return RedirectToAction("ResetPasswordConfirmation", "Account");
            }
            AddErrors(result);
            return View();
        }

        //
        // GET: /Account/ResetPasswordConfirmation
        [AllowAnonymous]
        public ActionResult ResetPasswordConfirmation()
        {
            return View();
        }

        //
        // POST: /Account/ExternalLogin
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult ExternalLogin(string provider, string returnUrl)
        {
            Session["Workaround"] = 0;
            // Request a redirect to the external login provider
            return new ChallengeResult(provider, Url.Action("ExternalLoginCallback", "Account", new { ReturnUrl = returnUrl }));
        }

        //
        // GET: /Account/SendCode
        [AllowAnonymous]
        public async Task<ActionResult> SendCode(string returnUrl, bool rememberMe)
        {
            var userId = await SignInManager.GetVerifiedUserIdAsync();
            if (userId == null)
            {
                return View("Error");
            }
            var userFactors = await UserManager.GetValidTwoFactorProvidersAsync(userId);
            var factorOptions = userFactors.Select(purpose => new SelectListItem { Text = purpose, Value = purpose }).ToList();
            return View(new SendCodeViewModel { Providers = factorOptions, ReturnUrl = returnUrl, RememberMe = rememberMe });
        }

        //
        // POST: /Account/SendCode
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> SendCode(SendCodeViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View();
            }

            // Generate the token and send it
            if (!await SignInManager.SendTwoFactorCodeAsync(model.SelectedProvider))
            {
                return View("Error");
            }
            return RedirectToAction("VerifyCode", new { Provider = model.SelectedProvider, ReturnUrl = model.ReturnUrl, RememberMe = model.RememberMe });
        }

        //
        // GET: /Account/ExternalLoginCallback
        [AllowAnonymous]
        public async Task<ActionResult> ExternalLoginCallback(string returnUrl)
        {
            var loginInfo = await AuthenticationManager.GetExternalLoginInfoAsync();

            if (loginInfo.Login.LoginProvider == "Google")
            {
                var externalIdentity = AuthenticationManager.GetExternalIdentityAsync(DefaultAuthenticationTypes.ExternalCookie);

                var emailClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email);
                var lastNameClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Surname);
                var givenNameClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.GivenName);
                //var addressClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.StreetAddress);
                //var countryClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Country);
                //var stateClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.StateOrProvince);
                //var postalClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.PostalCode);
                //var phoneClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.MobilePhone);
                //var genderClaim = externalIdentity.Result.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Gender);


                var email = emailClaim.Value;
                var firstName = givenNameClaim.Value;
                var lastname = lastNameClaim.Value;
                //var gender = genderClaim.Value;
                //var phone = phoneClaim.Value;

                string strConnectionSite = "";
                string strRequestURL = Request.Url.AbsoluteUri.ToString().ToLower();

                if (GlobalVariables.strSiteCode != null)
                {
                    strConnectionSite = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
                }
                else { strConnectionSite = "DefaultSiteDB"; }

                // Here check if user already exits and verfiyed or else show mail verification info page

                string strPassword = objDS.Encrypt("sp@123", "*!%$@~&#?,:");

                DataSet ds = new DataSet();
                string strQuery = "Select TOP 1 EM.EmpAutoID, EM.EmpLogin,EM.EmpMailId, EM.EmailIdentifier, EM.EmailVerify, SM.StatusDescription, EM.EmpStatus FROM JBM_Employeemaster EM INNER JOIN JBM_StatusMaster SM ON EM.EmailVerify=SM.StatusID WHERE EM.EmpMailId='" + email + "' or EM.EmpLogin='" + email + "' ORDER BY EM.EmpAutoID DESC";
                ds = DBProc.GetResultasDataSet(strQuery, strConnectionSite);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["EmailVerify"].ToString() == "1")
                    {
                        if (ds.Tables[0].Rows[0]["EmpStatus"].ToString() == "0")
                        {
                            return RedirectToAction("AccountRegistered", "Home", new { @id = "Blocked" });
                        }
                        else
                        {
                            System.Web.HttpContext.Current.Session["UserID"] = email;
                            return RedirectToAction("Dashboard", "Home");
                        }
                    }
                    else if (ds.Tables[0].Rows[0]["EmailVerify"].ToString() == "2")
                    {
                        return RedirectToAction("AccountRegistered", "Home", new { @id = ds.Tables[0].Rows[0]["EmailVerify"].ToString() });
                    }
                }
                else
                {
                    ds = new DataSet();
                    SqlConnection con = new SqlConnection();
                    con = DBProc.getConnection(strConnectionSite);
                    con.Open();
                    SqlCommand cmd = new SqlCommand("EmployeeRegister", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@EmpLogin", email);
                    cmd.Parameters.AddWithValue("@EmpPass", strPassword);
                    cmd.Parameters.AddWithValue("@EmpName", firstName);
                    cmd.Parameters.AddWithValue("@EmpSurname", lastname);
                    cmd.Parameters.AddWithValue("@DeptCode", 20);  // Default Content Inspection
                    cmd.Parameters.AddWithValue("@EmpMailId", email);
                    cmd.Parameters.AddWithValue("@RegistrType", "Google");
                    cmd.Parameters.AddWithValue("@EVerify", "2");
                    //cmd.ExecuteNonQuery();
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(ds);

                    con.Close();

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if (ds.Tables[0].Rows[0][0].ToString() == "Success")
                        {
                            string strMessage = "Hi " + firstName + ",<br/><br/>Thanks for signing up for Smart Publisher. You are just one step away from activation your account on the Smart Publisher.<br/><br/>Please verify the link <a href='https://localhost:44358/Home/AccountVerification/" + ds.Tables[0].Rows[0][4].ToString() + "?verify=" + ds.Tables[0].Rows[0][7].ToString().ToUpper() + "'>Activate Account</a><br/><br/>Regards,<br/>Smart Publisher<br/><br/>This is an auto generated mail.";
                            gen.Mail_Send("sivakumar.m@cenveo.com", "sivakumar.m@cenveo.com", "", "Smart Publisher Account Registration", strMessage);
                            return RedirectToAction("AccountRegistered", "Home", new { @id = "1" });
                        }
                        else
                        {
                            return RedirectToAction("Login", "Account");
                        }
                    }
                }

            }

            if (loginInfo == null)
            {
                return RedirectToAction("Login");
            }
            else {
                return RedirectToAction("AccountRegistered", "Home", new { @id = "Blocked" });
            }
            

            // Sign in the user with this external login provider if the user already has a login
            var result = await SignInManager.ExternalSignInAsync(loginInfo, isPersistent: false);
            switch (result)
            {
                case SignInStatus.Success:
                    return RedirectToLocal(returnUrl);
                case SignInStatus.LockedOut:
                    return View("Lockout");
                case SignInStatus.RequiresVerification:
                    return RedirectToAction("SendCode", new { ReturnUrl = returnUrl, RememberMe = false });
                case SignInStatus.Failure:
                default:
                    // If the user does not have an account, then prompt the user to create an account
                    ViewBag.ReturnUrl = returnUrl;
                    ViewBag.LoginProvider = loginInfo.Login.LoginProvider;
                    return View("ExternalLoginConfirmation", new ExternalLoginConfirmationViewModel { Email = loginInfo.Email });
            }
        }

        //
        // POST: /Account/ExternalLoginConfirmation
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> ExternalLoginConfirmation(ExternalLoginConfirmationViewModel model, string returnUrl)
        {
            if (User.Identity.IsAuthenticated)
            {
                return RedirectToAction("Index", "Manage");
            }

            if (ModelState.IsValid)
            {
                // Get the information about the user from the external login provider
                var info = await AuthenticationManager.GetExternalLoginInfoAsync();
                if (info == null)
                {
                    return View("ExternalLoginFailure");
                }
                var user = new ApplicationUser { UserName = model.Email, Email = model.Email };
                var result = await UserManager.CreateAsync(user);
                if (result.Succeeded)
                {
                    result = await UserManager.AddLoginAsync(user.Id, info.Login);
                    if (result.Succeeded)
                    {
                        await SignInManager.SignInAsync(user, isPersistent: false, rememberBrowser: false);
                        return RedirectToLocal(returnUrl);
                    }
                }
                AddErrors(result);
            }

            ViewBag.ReturnUrl = returnUrl;
            return View(model);
        }

        //
        // POST: /Account/LogOff
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LogOff()
        {
            AuthenticationManager.SignOut(DefaultAuthenticationTypes.ApplicationCookie);
            return RedirectToAction("Index", "Home");
        }

        //
        // GET: /Account/ExternalLoginFailure
        [AllowAnonymous]
        public ActionResult ExternalLoginFailure()
        {
            return View();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_userManager != null)
                {
                    _userManager.Dispose();
                    _userManager = null;
                }

                if (_signInManager != null)
                {
                    _signInManager.Dispose();
                    _signInManager = null;
                }
            }

            base.Dispose(disposing);
        }

        #region Helpers
        // Used for XSRF protection when adding external logins
        private const string XsrfKey = "XsrfId";

        private IAuthenticationManager AuthenticationManager
        {
            get
            {
                return HttpContext.GetOwinContext().Authentication;
            }
        }

        private void AddErrors(IdentityResult result)
        {
            foreach (var error in result.Errors)
            {
                ModelState.AddModelError("", error);
            }
        }

        private ActionResult RedirectToLocal(string returnUrl)
        {
            if (Url.IsLocalUrl(returnUrl))
            {
                return Redirect(returnUrl);
            }
            return RedirectToAction("Index", "Home");
        }

        [AllowAnonymous]
        [SessionExpire]
        public ActionResult ChangePassword()
        {
            return View();
        }
       

        internal class ChallengeResult : HttpUnauthorizedResult
        {
            public ChallengeResult(string provider, string redirectUri)
                : this(provider, redirectUri, null)
            {
            }

            public ChallengeResult(string provider, string redirectUri, string userId)
            {
                LoginProvider = provider;
                RedirectUri = redirectUri;
                UserId = userId;
            }

            public string LoginProvider { get; set; }
            public string RedirectUri { get; set; }
            public string UserId { get; set; }

            public override void ExecuteResult(ControllerContext context)
            {
                var properties = new AuthenticationProperties { RedirectUri = RedirectUri };
                if (UserId != null)
                {
                    properties.Dictionary[XsrfKey] = UserId;
                }
                context.HttpContext.GetOwinContext().Authentication.Challenge(properties, LoginProvider);
            }
        }
        #endregion
    }
}