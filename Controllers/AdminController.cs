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
using System.Collections.ObjectModel;
using System.Data.OleDb;

namespace SmartTrack.Controllers
{
    [SessionExpire]
    public class AdminController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        DataProc DBProc = new DataProc();
        // GET: Admin
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult AssignProject(string CustAcc, string EmpID, string SiteID)
        {
            if (CustAcc != null)
            {
                Session["UserID"] = EmpID;
                Session["sCustAcc"] = CustAcc;
                Session["sSiteID"] = SiteID;


                clsCollec.getSiteDBConnection(SiteID, CustAcc);
            }
            GlobalVariables.strCustAcc = Session["sCustAcc"].ToString();
            DataSet ds = new DataSet();

            DataSet dscustlst = new DataSet();
            dscustlst = DBProc.GetResultasDataSet("SELECT distinct JBM_Info.CustID,BK_ProcessInfo.EmpAutoID FROM  JBM_Info INNER JOIN  BK_ProcessInfo ON JBM_Info.JBM_AutoID = BK_ProcessInfo.JBM_AutoID where BK_ProcessInfo.AccTime =(select max(AccTime) from BK_ProcessInfo)", Session["sConnSiteDB"].ToString());
           
            //Load Customer list Items
            List<Chkcust> Customeritems = new List<Chkcust>();
            //List<SelectListItem> Customeritems = new List<SelectListItem>();
            ds = new DataSet();
            ds = DBProc.GetResultasDataSet("select CustID,CustName from jbm_customermaster", Session["sConnSiteDB"].ToString());                     

            if (dscustlst.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ViewBag.EmpAutoID = dscustlst.Tables[0].Rows[0]["EmpAutoID"].ToString();
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        bool ischk = false;
                        foreach (DataRow myRowcust in dscustlst.Tables[0].Rows)
                        {                           
                            if (myRow["CustID"].ToString()== myRowcust["CustID"].ToString())
                            {
                                ischk = true;
                            }
                           
                        }
                        Customeritems.Add(new Chkcust
                        {
                            Text = myRow["CustName"].ToString(),
                            Value = myRow["CustID"].ToString(),
                            IsChecked = ischk
                        });
                    }
                }
            }
            else
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow myRow in ds.Tables[0].Rows)
                    {
                        Customeritems.Add(new Chkcust
                        {
                            Text = myRow["CustName"].ToString(),
                            Value = myRow["CustID"].ToString(),
                            IsChecked = false
                        });
                    }
                }
            }
            ViewBag.Customerlist = Customeritems;

            //Load User list Items
            var objCollec = new clsCollection();
            ViewBag.Userlist = objCollec.GetUserCollection(Session["sConnSiteDB"].ToString());
            return View();
        }
        public JsonResult GetProjectlst(List<Chklst> AssignPrjlst)
        {
            try
            {
                List<Chkcust> Projectitems = new List<Chkcust>();
                //List<SelectListItem> Projectitems = new List<SelectListItem>();
                int count = AssignPrjlst.Count;
                for (int i = 0; i < count; i++)
                {
                    string custid = AssignPrjlst[i].Chkid.ToString().Trim();
                    string Eid ="";
                    if (AssignPrjlst[i].Empid != null)
                    {
                        Eid = AssignPrjlst[i].Empid.ToString().Trim();
                    }
                    if (custid.Trim() != "")
                    {
                        //Load Project list Items based on customer
                        DataSet dsprjlst = new DataSet();
                        dsprjlst = DBProc.GetResultasDataSet("SELECT distinct JBM_Info.JBM_AutoID FROM  JBM_Info INNER JOIN  BK_ProcessInfo ON JBM_Info.JBM_AutoID = BK_ProcessInfo.JBM_AutoID where BK_ProcessInfo.EmpAutoID = '" + Eid + "'", Session["sConnSiteDB"].ToString());

                        DataSet ds = new DataSet();
                        ds = DBProc.GetResultasDataSet("select JBM_AutoID,JBM_ID from JBM_Info where CustID='"+ custid.Trim() + "'", Session["sConnSiteDB"].ToString());
                        if (dsprjlst.Tables[0].Rows.Count > 0)
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                foreach (DataRow myRow in ds.Tables[0].Rows)
                                {
                                    bool ischk = false;
                                    foreach (DataRow myRowcust in dsprjlst.Tables[0].Rows)
                                    {
                                        if (myRow["JBM_AutoID"].ToString() == myRowcust["JBM_AutoID"].ToString())
                                        {
                                            ischk = true;
                                        }

                                    }
                                    Projectitems.Add(new Chkcust
                                    {
                                        Text = myRow["JBM_ID"].ToString(),
                                        Value = myRow["JBM_AutoID"].ToString(),
                                        IsChecked = ischk
                                    });
                                }
                            }
                        }
                        else
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                foreach (DataRow myRow in ds.Tables[0].Rows)
                                {
                                    Projectitems.Add(new Chkcust
                                    {
                                        Text = myRow["JBM_ID"].ToString(),
                                        Value = myRow["JBM_AutoID"].ToString(),
                                        IsChecked = false
                                    });
                                }
                            }
                        }
                        ViewBag.Projectlist = Projectitems;
                    }
                   
                }
                return Json(Projectitems, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult GetCustProjectlst(string Empid)
        {
            try
            {
                List<Chkcust> Customeritems = new List<Chkcust>();
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("select CustID,CustName from jbm_customermaster", Session["sConnSiteDB"].ToString());
                DataSet dscustlst = new DataSet();
                dscustlst = DBProc.GetResultasDataSet("SELECT distinct BK_ProcessInfo.AutoSeqID, JBM_Info.CustID,JBM_Info.JBM_AutoID FROM  JBM_Info INNER JOIN  BK_ProcessInfo ON JBM_Info.JBM_AutoID = BK_ProcessInfo.JBM_AutoID where  BK_ProcessInfo.EmpAutoID = '"+Empid+"'", Session["sConnSiteDB"].ToString());
                if (dscustlst.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            bool ischk = false;
                            foreach (DataRow myRowcust in dscustlst.Tables[0].Rows)
                            {
                                if (myRow["CustID"].ToString() == myRowcust["CustID"].ToString())
                                {
                                    ischk = true;
                                }

                            }
                            Customeritems.Add(new Chkcust
                            {
                                Text = myRow["CustName"].ToString(),
                                Value = myRow["CustID"].ToString(),
                                IsChecked = ischk
                            });
                        }
                    }
                }
                else
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow myRow in ds.Tables[0].Rows)
                        {
                            Customeritems.Add(new Chkcust
                            {
                                Text = myRow["CustName"].ToString(),
                                Value = myRow["CustID"].ToString(),
                                IsChecked = false
                            });
                        }
                    }
                }
                return Json(Customeritems, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult SaveProject(List<custprj> custprj)
        {
            try
            {
                SqlConnection con = new SqlConnection();              
               
                con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                con.Open();
                int count = custprj[0].Prjlst.Count;
                int countcust= custprj.Count;
                string qry="";
                string eid = "";
                for (int i = 0; i < count; i++)
                {
                    string prjid;
                    if (custprj[0].Prjlst[i].Chkprjid == null)
                    {
                        string custid = custprj[i].Custid.ToString().Trim();
                        eid = custprj[i].Prjlst[i].Empid.ToString().Trim();
                        DataSet dslst1 = DBProc.GetResultasDataSet("SELECT distinct BK_ProcessInfo.AutoSeqID, JBM_Info.CustID,JBM_Info.JBM_AutoID FROM  JBM_Info INNER JOIN  BK_ProcessInfo ON JBM_Info.JBM_AutoID = BK_ProcessInfo.JBM_AutoID WHERE  BK_ProcessInfo.EmpAutoID = '" + eid + "' and JBM_Info.CustID = '" + custid + "'", Session["sConnSiteDB"].ToString());
                        if (dslst1.Tables[0].Rows.Count > 0)
                        {
                            SqlCommand cmd = new SqlCommand();
                            
                            foreach (DataRow myRow in dslst1.Tables[0].Rows)
                            {
                                cmd = new SqlCommand("delete from " + GlobalVariables.strCustAcc + "_ProcessInfo  where AutoSeqID='" + myRow["AutoSeqID"].ToString() + "'", con);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        return Json(new { dataSch = "NotInPrj" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        prjid = custprj[0].Prjlst[i].Chkprjid.ToString().Trim();
                    }
                    
                    string AccTime = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                    string empid;
                    if (custprj[0].Prjlst[i].Empid == null)
                    {
                        return Json(new { dataSch = "NotIn" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        empid = custprj[0].Prjlst[i].Empid.ToString().Trim();
                    }
                    if (prjid.Trim() != "")
                    {
                        qry += "'" + prjid + "',";
                        eid = empid;
                        DataSet dscustlst = new DataSet();
                        dscustlst = DBProc.GetResultasDataSet("select * from BK_ProcessInfo where EmpAutoID='" + empid + "' and JBM_AutoID='" + prjid + "'", Session["sConnSiteDB"].ToString());
                        SqlCommand cmd = new SqlCommand();
                        //cmd = new SqlCommand("delete from BK_ProcessInfo", con);
                        //cmd.ExecuteNonQuery();
                        if (dscustlst.Tables[0].Rows.Count > 0)
                        {
                            cmd = new SqlCommand("update " + GlobalVariables.strCustAcc + "_ProcessInfo set EmpAutoID='" + empid + "',JBM_AutoID='" + prjid + "' where AutoSeqID='" + dscustlst.Tables[0].Rows[0][0].ToString() + "'", con);
                        }
                        else
                        {
                            cmd = new SqlCommand("insert into " + GlobalVariables.strCustAcc + "_ProcessInfo(EmpAutoID,JBM_AutoID,AccTime,ProcessID,Description) Values('" + empid + "','" + prjid + "','" + AccTime + "','0','')", con);
                        }
                        cmd.ExecuteNonQuery();
                    }

                }
                qry = qry.Remove(qry.Length - 1, 1);
                DataSet dslst = new DataSet();
                for (int i = 0; i < countcust; i++)
                {
                    string custid = custprj[i].Custid.ToString().Trim();
                    dslst = DBProc.GetResultasDataSet("SELECT distinct BK_ProcessInfo.AutoSeqID, JBM_Info.CustID,JBM_Info.JBM_AutoID FROM  JBM_Info INNER JOIN  BK_ProcessInfo ON JBM_Info.JBM_AutoID = BK_ProcessInfo.JBM_AutoID WHERE JBM_Info.JBM_AutoID Not IN(" + qry + ") and BK_ProcessInfo.EmpAutoID = '" + eid + "' and JBM_Info.CustID = '" + custid + "'", Session["sConnSiteDB"].ToString());
                    if (dslst.Tables[0].Rows.Count > 0)
                    {
                        SqlCommand cmd = new SqlCommand();
                        foreach (DataRow myRow in dslst.Tables[0].Rows)
                        {
                            cmd = new SqlCommand("delete from " + GlobalVariables.strCustAcc + "_ProcessInfo  where AutoSeqID='" + myRow["AutoSeqID"].ToString() + "'", con);
                            cmd.ExecuteNonQuery();
                        }
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
        
        public ActionResult MenuConfig(string CustAcc, string EmpID, string SiteID)
        {
            try
            {

                Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
              
                var objCollec = new clsCollection();
                ViewBag.Userlist = objCollec.GetUserCollection(Session["sConnSiteDB"].ToString());
                ViewBag.Departmentlist = objCollec.GetDepartments(Session["sConnSiteDB"].ToString());

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select MenuID,MenuDispName from JBM_MenuMaster", Session["sConnSiteDB"].ToString());

                List<SelectListItem> collmenuList = ds.Tables[0].AsEnumerable()
                                                 .Select(dataRow => new SelectListItem
                                                 {
                                                     Value = dataRow.Field<int>("MenuID").ToString(),
                                                     Text = dataRow.Field<string>("MenuDispName")
                                                 }).ToList();

                ViewBag.MenuList = collmenuList;
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select distinct CustType from JBM_Customermaster", Session["sConnSiteDB"].ToString());

                List<SelectListItem> collCustTypeList = ds.Tables[0].AsEnumerable()
                                                 .Select(dataRow => new SelectListItem
                                                 {
                                                     Value = dataRow.Field<string>("CustType").ToString(),
                                                     Text = dataRow.Field<string>("CustType")
                                                 }).ToList();

                ViewBag.CustTypeList = collCustTypeList;

                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select MenuID,MenuDispName from JBM_MenuMaster WHERE MenuType='MenuItem'", Session["sConnSiteDB"].ToString());

                List<SelectListItem> collMainMenuList = ds.Tables[0].AsEnumerable()
                                                 .Select(dataRow => new SelectListItem
                                                 {
                                                     Value = dataRow.Field<int>("MenuID").ToString(),
                                                     Text = dataRow.Field<string>("MenuDispName")
                                                 }).ToList();

                ViewBag.MainMenuList = collMainMenuList;

                return View();
            }
            catch (Exception)
            {

                throw;
            }
            
        }
        
        public ActionResult GetMenuList()
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
        
        public ActionResult AddNewMenu(string objData)
        {
            try
           {

                if (Session["sConnSiteDB"].ToString() == "")
                {
                    Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
                }

                string s = System.Uri.UnescapeDataString(objData);
                List<string> saveIds = JsonConvert.DeserializeObject<List<string>>(s);
                if (saveIds.Count > 0)
                {
                    for (int i = 0; i < saveIds.Count; i++)
                    {
                        string strMenuID = "";  string strMenuDisplay = ""; string strMenuIcon = ""; string strMenuLink = ""; string strChecked = ""; string strMenuType = ""; string strCustGroup = "";
                        strMenuID = saveIds[i].Split('|')[0];
                        if (string.IsNullOrEmpty(strMenuID) || strMenuID == "undefined") { strMenuID = ""; }

                        strMenuDisplay = saveIds[i].Split('|')[1];
                        strMenuLink = saveIds[i].Split('|')[2];
                        strMenuIcon = saveIds[i].Split('|')[3];
                        strChecked = saveIds[i].Split('|')[4];
                        strMenuType = saveIds[i].Split('|')[5];
                        strCustGroup = saveIds[i].Split('|')[6];
                        strCustGroup = strCustGroup.Replace(",", "|");

                        //To Insert Into JBM_MenuMaster
                        string strResult = "";
                        if (strMenuType == "MenuItem")
                        {
                            strResult = DBProc.GetResultasString("Insert Into JBM_MenuMaster (MenuID,SubMenuUnder,MenuType,MenuIcon,MenuDispName,MenuAction,CustomerGroup,Flag) Values (isnull((Select max(MenuID) + 1 from JBM_MenuMaster),1), isnull((Select max(MenuID) + 1 from JBM_MenuMaster),1), '" + strMenuType + "', '" + strMenuIcon + "', '" + strMenuDisplay + "', '" + strMenuLink + "', '" + strCustGroup + "', '" + strChecked + "')", Session["sConnSiteDB"].ToString());
                        }
                        else {
                            strResult = DBProc.GetResultasString("Insert Into JBM_MenuMaster (MenuID,SubMenuUnder,MenuType,MenuIcon,MenuDispName,MenuAction,CustomerGroup,Flag) Values (isnull((Select max(MenuID) + 1 from JBM_MenuMaster),1),'" + strMenuID+ "','" + strMenuType + "', '" + strMenuIcon + "', '" + strMenuDisplay + "', '" + strMenuLink + "', '" + strCustGroup + "', '" + strChecked + "')", Session["sConnSiteDB"].ToString());
                        }


                    }

                }
                else
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }


        /// <summary>
        /// User Settings
        /// </summary>
        /// <returns></returns>
        
        public ActionResult UserSettings(string CustAcc, string EmpID, string SiteID)
        {
            clsCollec.getSiteDBConnection(SiteID, CustAcc);
            if (Session["sConnSiteDB"].ToString() == "")
            {
                Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString();
            }

            return View();
        }

        public NASViewModel selectDataEMP(string vDeptCode, string vEmpList)
        {
            try
            {

                NASViewModel model = new NASViewModel();

                string qryJoin = "";
                if (vEmpList != "")
                {
                    vEmpList = vEmpList.Replace(",", "','");
                    qryJoin = "em.EmpAutoID in ('" + vEmpList + "')";
                }
                else {
                    qryJoin = "em.DeptCode='" + vDeptCode + "'";
                }

                string strQueryFinal = "Select EmpAutoID,EmpLogin,EmpName,dm.DeptName,SiteID,NASAccessType,NASAccessCmd,EmpStatus from JBM_EmployeeMaster em inner join JBM_Departmentmaster dm on dm.DeptCode=em.DeptCode WHERE " + qryJoin + " and EmpStatus='1'";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(strQueryFinal + "  order by EmpName asc ", Session["sConnSiteDB"].ToString());

                foreach (DataRow row in dt.Rows)
                {
                    
                    string profilelink = profileImgLink(Convert.ToString(row["EmpLogin"]), "/Admin/getNASAccess");
                    profilelink = profilelink.Replace("?vDeptCode=" + vDeptCode, "");
                    profilelink = profilelink.Replace("/Admin/NASRights?CustAcc=" + Session["sCustAcc"].ToString() + "&EmpID=" + Session["UserID"].ToString() + "&SiteID=" + Session["sSiteID"].ToString(), "");
                    profilelink = profilelink.ToLower().Replace("/admin/nasrights", "");
                    profilelink = profilelink.ToLower().Replace("/admin/getnasaccess", "");
                    
                    NAS fld = new NAS();
                    fld.EmpAutoID = Convert.ToString(row["EmpAutoID"]);
                    fld.EmpLogin = Convert.ToString(row["EmpLogin"]);
                    fld.EmpName = Convert.ToString(row["EmpName"]);
                    fld.DeptName = Convert.ToString(row["DeptName"]);
                    fld.NASAccessCmd = Convert.ToString(row["NASAccessCmd"]);
                    fld.NASAccessType = Convert.ToString(row["NASAccessType"]);
                    fld.Profilelink = profilelink;
                    fld.EmpStatus = Convert.ToString(row["EmpStatus"]);
                    model.dt.Add(fld);
                }
                return model;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public PartialViewResult getNASAccess(string vDeptCode, string vEmpList)
        {
            NASViewModel model = new NASViewModel();
            model = selectDataEMP(vDeptCode, vEmpList);
            return PartialView("_EmpTestPartial", model);
            //return View(model);

            //try
            //{
            //    clsCollec.getSiteDBConnection("L0001", "BK");
            //    if (Session["sConnSiteDB"].ToString() == "")
            //    {
            //        Session["sConnSiteDB"] = GlobalVariables.strConnSite;
            //    }

            //    string strQueryFinal = "Select EmpAutoID,EmpLogin,EmpName,dm.DeptName,SiteID,NASAccessType,NASAccessCmd,EmpStatus from JBM_EmployeeMaster em inner join JBM_Departmentmaster dm on dm.DeptCode=em.DeptCode";

            //    DataSet ds = new DataSet();
            //    ds = DBProc.GetResultasDataSet(strQueryFinal + "  order by EmpName asc ", Session["sConnSiteDB"].ToString());

            //    var JSONString = from a in ds.Tables[0].AsEnumerable()
            //                     select new[] {
            //                         CreateDynamicItem(a[0].ToString(),a[1].ToString(),a[2].ToString(),a[3].ToString(),a[5].ToString(),a[6].ToString(),a[7].ToString(), "UserInfo"),
            //                         a[0].ToString() + "#" + a[3].ToString() + "#" + a[6].ToString(),
            //                         a[0].ToString() + "#" + a[3].ToString() + "#" + a[6].ToString(),
            //                         CreateDynamicItem(a[0].ToString(),a[1].ToString(),a[2].ToString(),a[3].ToString(),a[5].ToString(),a[6].ToString(),a[7].ToString(), "Action")
            //     };
            //    return Json(new { dataResult = JSONString }, JsonRequestBehavior.AllowGet);
            //}
            //catch (Exception)
            //{
            //    return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            //}
        }

        public string CreateDynamicItem(string EmpAutoID, string EmpLogin, string EmpName, string DeptName, string NASAccssType, string NASCmd,string Status, string strType)
        {
            string formControl = string.Empty;
            try  //onkeypress='funcInputValidate()'
            {
                if (strType == "UserInfo")
                {
                   string profilelink  =  profileImgLink(EmpLogin, "/Admin/getNASAccess");
                    string strStatus = "<span class='badge bg-danger'>In-Active</span>";
                    if (Status == "1")
                    {
                        strStatus = "<span class='badge bg-success'>Active</span>";
                    }

                    formControl = "<div class='container'><div class='row'><div style='width: 35%;'><img class='user-image' src='" + profilelink + "' style='border-radius: 50%;width: 50%;'></div><div style='width: 65%;'><span><b>" + EmpName + ",</b></span><br><span>" + DeptName  + "</span><br><span>Login ID: " + EmpLogin + "</span><br>" + strStatus + "</div></div></div>";
                }
                else if (strType == "CeninwAccess")
                {
//                    formControl = "<select id='ddlEmpStatus' class='form-control select2 select2-danger' data-dropdown-css-class='select2-danger' style='width: 100%;'><option value = '1' selected = '' >[Choose the Ceninw Access] </option><option value = '1'>Read</option><option value = '0'>Partial Access</option><option value = '0'>Full Access</option></select><div class='container'><div class='row'><div class='form-check' style=' width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckChecked' checked=''><label class='form-check-label' for='flexCheckChecked'>Create</label></div><div class='form-check' style='width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckDefault'><label class='form-check-label' for='flexCheckDefault'>Delete</label></div><div class='form-check' style='width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckDefault'><label class='form-check-label' for='flexCheckDefault'>Copy</label></div><div class='form-check' style='width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckDefault'><label class='form-check-label' for='flexCheckDefault'>Paste</label></div><div class='form-check' style='width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckDefault'><label class='form-check-label' for='flexCheckDefault'>Upload</label></div><div class='form-check' style='width: 25%;'><input class='form-check-input' type='checkbox' value='' id='flexCheckDefault'><label class='form-check-label' for='flexCheckDefault'>Rename</label></div></div></div>";
                }
                else if (strType == "CenproAccess")
                {
  //                  formControl = "";
                }
              
                else if (strType == "Action")
                {
                    formControl = "<button type='button' id='Update" + EmpAutoID + "' onClick=\"funcUserNASEdit('" + EmpAutoID + "')\" class='btn btn-sm btn-outline-danger' style='width:100%' name='update' value='Update'><span class='fas fa-save fa-1x text-grey'></span>&nbsp;Update</button>";

                }


                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }

        public string profileImgLink(string imguserId, string strReplURL)
        {

            string strUrl = Request.Url.AbsoluteUri.ToString(); //.Replace(Request.RawUrl.ToString(), "")
            //strUrl = strUrl.Substring(0, strUrl.IndexOf('?'));
            //string strUrl = Request.Url.AbsoluteUri.ToString().Replace(strReplURL, "").Replace("?vStatus=undefined", "");
            try
            {
                string strPath = System.Web.Hosting.HostingEnvironment.MapPath(@"~/Images/Employee/" + imguserId + ".png");

                if (System.IO.File.Exists(strPath))
                {
                    return strUrl + "/Images/Employee/" + imguserId + ".png";
                }
                else
                {
                    strPath = strPath.Replace(".png", ".jpg");
                    if (System.IO.File.Exists(strPath))
                    {
                        return strUrl + "/Images/Employee/" + imguserId + ".jpg" ;
                    }
                    else
                    {
                        strPath = strPath.Replace(".jpg", ".gif");
                        if (System.IO.File.Exists(strPath))
                        {
                            return strUrl + "/Images/Employee/" + imguserId + ".gif";
                        }
                        return strUrl + "/Images/Employee/profile.png" ;
                    }

                }
                
            }
            catch (Exception)
            {
                return strUrl + "/Images/Employee/" + imguserId + ".gif";
            }
        }
        
        public ActionResult AddNewUser()
        {
            return View();
        }
        
        public ActionResult NASRights()//(string CustAcc, string EmpID, string SiteID)
        {
            try
            {

                string strUrl = Request.Url.AbsolutePath.ToString();
                Session["returnURL"] = strUrl; // "http://10.20.11.31/smarttrack/ManagerInbox.aspx";

                //if (CustAcc != null)
                //{
                //    Session["sCustAcc"] = CustAcc;
                //    Session["sSiteID"] = SiteID;
                //    Session["sConnSiteDB"] = "";
                //    Session["UserID"] = EmpID;

                //    clsCollec.getSiteDBConnection(SiteID, CustAcc);
                //}

                //if (Session["sConnSiteDB"].ToString() == "")
                //{
                //    Session["sConnSiteDB"] = "dbConnSmart" + GlobalVariables.strSiteCode.Trim() + "-" + GlobalVariables.strEnvironment.Trim().ToString(); 
                //}

                //DataSet ds = new DataSet();    
                //ds = DBProc.GetResultasDataSet("Select a.EmpAutoid, a.EmpLogin, a.EmpPass, a.EmpName, a.EmpLoginName,a.EmpMailId,a.RoleID, a.DeptCode,a.DeptAccess, a.TeamPlayer,(Select b.DeptName from JBM_DepartmentMaster b  where a.DeptCode = b.DeptCode) as DeptName, a.CustAccess, a.TeamMasterAccDept, a.GroupMenu, a.JwAccessItm, a.BMAccessItm, a.TeamID, a.SiteID, a.SubTeam, a.QecTeamID, a.EmpSurname,a.etype,a.empqc,a.DesignationCode,a.TLEmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.TLEmpAutoID) as TLEmpName ,a.MGREmpAutoID ,(select EmpName from JBM_EmployeeMaster where EmpAutoID = a.MGREmpAutoID) as MGREmpName,a.SiteAcc, a.ProfilePassword,a.Ven_Site,a.ServiceTaxno from JBM_EmployeeMaster a WHERE a.Emplogin = '" + EmpID + "' and a.EmpStatus = '1'", Session["sConnSiteDB"].ToString());
                //if (ds.Tables[0].Rows.Count > 0)
                //{
                //    Session["EmpAutoId"] = ds.Tables[0].Rows[0]["EmpAutoId"].ToString();
                //    Session["EmpName"] = ds.Tables[0].Rows[0]["EmpName"].ToString();

                //    Session["DeptName"] = ds.Tables[0].Rows[0]["DeptName"].ToString();
                //    Session["DeptCode"] = ds.Tables[0].Rows[0]["DeptCode"].ToString();
                //    Session["RoleID"] = ds.Tables[0].Rows[0]["RoleID"].ToString();
                //    Session["UserID"] = ds.Tables[0].Rows[0]["EmpLogin"].ToString();
                //    Session["AccessRights"] = "";
                //}

                DataSet ds = new DataSet();
                List<SelectListItem> lstDept = new List<SelectListItem>();

                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select DeptCode,DeptName from JBM_DepartmentMaster WHERE DeptCode in (select distinct DeptCode from JBM_Employeemaster where EmpStatus=1)", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        lstDept.Add(new SelectListItem
                        {
                            Text = ds.Tables[0].Rows[intCount]["DeptName"].ToString(),
                            Value = ds.Tables[0].Rows[intCount]["DeptCode"].ToString()
                        });
                    }
                }
                ViewBag.Deptlist = lstDept;

                List<SelectListItem> lstEmp = new List<SelectListItem>();
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select EmpAutoid,EmpLogin, EmpName from JBM_EmployeeMaster WHERE  EmpStatus=1", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        lstEmp.Add(new SelectListItem
                        {
                            Text = ds.Tables[0].Rows[intCount]["EmpName"].ToString() + " (" + ds.Tables[0].Rows[intCount]["EmpLogin"].ToString() + ")",
                            Value = ds.Tables[0].Rows[intCount]["EmpAutoid"].ToString()
                        });
                    }
                }
                ViewBag.Emplist = lstEmp;

                NASViewModel model = new NASViewModel();
                model = selectDataEMP("180", "");


                //string strQueryFinal = "Select top 20 EmpAutoID,EmpLogin,EmpName,dm.DeptName,SiteID,NASAccessType,NASAccessCmd,EmpStatus from JBM_EmployeeMaster em inner join JBM_Departmentmaster dm on dm.DeptCode=em.DeptCode";

                //DataTable dt = new DataTable();
                //dt = DBProc.GetResultasDataTbl(strQueryFinal + "  order by EmpName asc ", Session["sConnSiteDB"].ToString());

                //foreach (DataRow row in dt.Rows)
                //{
                //    string profilelink = profileImgLink(Convert.ToString(row["EmpLogin"]), "/Admin/NASAccess");

                //    NAS fld = new NAS();
                //    fld.EmpAutoID = Convert.ToString(row["EmpAutoID"]);
                //    fld.EmpLogin = Convert.ToString(row["EmpLogin"]);
                //    fld.EmpName = Convert.ToString(row["EmpName"]);
                //    fld.DeptName = Convert.ToString(row["DeptName"]);
                //    fld.NASAccessCmd = Convert.ToString(row["NASAccessCmd"]);
                //    fld.Profilelink = profilelink;
                //    fld.EmpStatus = Convert.ToString(row["EmpStatus"]);
                //    model.dt.Add(fld);
                //}
                //return PartialView("_EmpTestPartial", model);
                return View(model);
                //return View();
            }
            catch (Exception)
            {
                return View();
            }

        }
       
        public ActionResult UpdateNASAccessDetails(string sEmpAutoID, string sNASType, string sNASCmd)
        {
            try
            {
                string result = DBProc.GetResultasString("UPDATE JBM_EmployeeMaster SET NASAccessType='" + sNASType + "', NASAccessCmd='" + sNASCmd + "' WHERE  EmpAutoID='" + sEmpAutoID + "'", Session["sConnSiteDB"].ToString());

                if (result == "0")
                {
                    return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataSch = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
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