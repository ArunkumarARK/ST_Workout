using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Web;
using System.Web.Mvc;
using SmartTrack.Helper;
using SmartTrack.Models;
using System.Text;
using System.Net.Http;
using System.Web.UI.WebControls;
using Newtonsoft.Json;
using System.Net;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Web.Http.Results;
using DocumentFormat.OpenXml;
using System.IO;
using Microsoft.VisualBasic.ApplicationServices;
using System.Xml.Linq;

namespace SmartTrack.Controllers
{
    //[RequireHttps]
    public class HomeController : Controller
    {
        DataProc DBProc = new DataProc();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        //[Authorize]
        [SessionExpire]
        public ActionResult Dashboard()
        {
            //string strQuery = "Select * from JBM_MenuMaster where MenuID in (select * from dbo.CSVToTable((select CONVERT(varchar(25), replace(MenuIDs,'|',',')) from JBM_EmployeeMenu where EmpAutoID='" + Session["EmpAutoId"].ToString()  + "'))) and Flag=1";
            string strQuery = "Select * from JBM_MenuMaster where MenuID in (select * from dbo.CSVToTable((select CONVERT(varchar(25), replace((Select CASE When GroupName is not null then (select MenuIDs from JBM_GroupMenu where GroupName=(Select  GroupName  from JBM_EmployeeMenu JE  WHERE EmpAutoID='" + Session["EmpAutoId"].ToString() + "')) Else (Select  MenuIDs  from JBM_EmployeeMenu JE  WHERE EmpAutoID='" + Session["EmpAutoId"].ToString() + "') END as  MenuIDs from JBM_EmployeeMenu WHERE EmpAutoID='" + Session["EmpAutoId"].ToString() + "'),'|',',')) from JBM_EmployeeMenu where EmpAutoID='" + Session["EmpAutoId"].ToString() + "'))) and Flag=1 and (CustomerGroup like  '%BK%' or CustomerGroup is null or CustomerGroup = '')";
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                //TempData["dept"] = ds.Tables[0].Rows[0][6].ToString();
                Session["MenuMaster"] = ds.Tables[0];

                List<MenuBindModels> _menuList = ds.Tables[0].AsEnumerable().Select(
                            dataRow => new MenuBindModels
                            {
                                mMainMenuId = dataRow.Field<int>("MenuID"),
                                mMenuDispName = dataRow.Field<string>("MenuDispName"),
                                mSubMenuId = dataRow.Field<int>("SubMenuUnder"),
                                mMenuType = dataRow.Field<string>("MenuType"),
                                mMenuIcon = dataRow.Field<string>("MenuIcon"),
                                mMenuAction = dataRow.Field<string>("MenuAction")
                            }).ToList();
                Session["MenuMaster"] = _menuList;
                //return RedirectToAction("Dashboard", "Home");
                return View();


            }

            return View();
        }


        [AllowAnonymous]
        public ActionResult AccountVerification(string id, string verify)
        {
            try
            {
                string strQuery = "Update JBM_EmployeeMaster SET EmailVerify='1' Output Inserted.EmailVerify WHERE EmpAutoID='" + id + "' and EmailIdentifier='" + verify.ToUpper() + "'";
                DataSet strResult = new DataSet();
                strResult = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());

                if (strResult.Tables[0].Rows.Count > 0)
                {
                    if (strResult.Tables[0].Rows[0][0].ToString() == "1")
                    {
                        ViewBag.StatusText = "IsVerified";
                    }
                    else
                    {
                        ViewBag.StatusText = "InValid";
                    }
                }
                else
                {
                    ViewBag.StatusText = "InValid";
                }


                return View();
            }
            catch (Exception ex)
            {
                ViewBag.StatusText = ex.Message;
                return View();
            }

        }
        [AllowAnonymous]
        public ActionResult AccountRegistered(string id)
        {
            try
            {
                if (id == "1")
                {
                    ViewBag.StatusText = "IsRegister";
                }
                else if (id == "2")
                {
                    ViewBag.StatusText = "Exists";
                }
                else if (id == "Blocked")
                {
                    ViewBag.StatusText = "Blocked";
                }
                return View();
            }
            catch (Exception ex)
            {
                ViewBag.StatusText = ex.Message;
                return View();
            }
        }

        [SessionExpire]
        public ActionResult ChangePasswordUpdate(string vCurrentPassword, string vNewPassword)
        {
            try
            {
                if (Session["sConnSiteDB"].ToString() == "")
                {
                    Session["sConnSiteDB"] = GlobalVariables.strConnSite;
                }

                string strPassword = ASCIItoHex(vNewPassword.ToString());
                string strCurrPassword = ASCIItoHex(vCurrentPassword.ToString());

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select EmpPass from JBM_EmployeeMaster  WHERE EmpLogin='" + Session["EmpIdLogin"].ToString() + "'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    string strCurrPassDB = ds.Tables[0].Rows[0]["EmpPass"].ToString();
                    if (strCurrPassword.Trim().ToUpper() == strCurrPassDB.Trim().ToUpper())
                    {
                        string strStatus = DBProc.GetResultasString("UPDATE JBM_EmployeeMaster SET EmpPass='" + strPassword + "' WHERE EmpLogin='" + Session["EmpIdLogin"].ToString() + "'", Session["sConnSiteDB"].ToString());

                        return Json(new { dataRes = "Success" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { dataRes = "CurrentPassNotMatch" }, JsonRequestBehavior.AllowGet);
                    }
                }

                return Json(new { dataRes = "CurrentPassNotMatch" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataRes = "Failed" }, JsonRequestBehavior.AllowGet);
            }
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
        [AllowAnonymous]
        public ActionResult EmeraldAPIUpload()
        {
            return View();
        }
        public ActionResult EmeraldProjectVerify(string projectId)
        {
            try
            {
                string strQuery = "Select JBM_AutoID,JBM_ID from JBM_Info where jbm_id='" + projectId + "'"; 
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQuery, "dbConnSmartCH-LIVE");
                if (ds.Tables[0].Rows.Count == 0)
                {
                    return Json(new { aaData = "Project ID not match" }, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception ex)
            {
                return Json(new { aaData = ex.Message }, JsonRequestBehavior.AllowGet);
            }
            return Json(new { aaData = "Project ID Matched" }, JsonRequestBehavior.AllowGet);

        }
        [HttpPost]
        public ActionResult Index(List<HttpPostedFileBase> file, string projectid)
        {
            string strAutoArtID = string.Empty;
            string strJBMAID = string.Empty;
            string strProjectID = string.Empty;

            string strQuery = "Select Top 1 j.JBM_AutoID,j.JBM_ID,b.AutoArtID from JBM_Info j inner join BK_ChapterInfo b on j.JBM_AutoID=b.JBM_AutoID where j.jbm_id='" + projectid + "'";
            DataSet ds = new DataSet();
            ds = DBProc.GetResultasDataSet(strQuery, "dbConnSmartCH-LIVE");
            if (ds.Tables[0].Rows.Count == 0)
            {
                return Json(new { dataUp = "Project ID not match" }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                strAutoArtID = ds.Tables[0].Rows[0]["AutoArtID"].ToString();
                strJBMAID = ds.Tables[0].Rows[0]["JBM_AutoID"].ToString();
                strProjectID = ds.Tables[0].Rows[0]["JBM_ID"].ToString();
            }

            string strURL = "https://ecms.emeraldgroup.com/alfresco/api/-default-/public/alfresco/versions/1/nodes/390ef4eb-7112-48c9-87e8-4bf7049fae25/children";
            string strUserID = "KGLBooksDelivery";
            string strPwd = "SvYrZMw5sQN7xtuYh4Tk5hTEFKY36xp5xCQGWx8L2FHKUtZzXQsdT8A7pg8mgKwP";
            string strTicket = "";

            var data = new { userId = strUserID, password = strPwd };
            string json = JsonConvert.SerializeObject(data);
            string url = "https://ecms.emeraldgroup.com/alfresco/api/-default-/public/authentication/versions/1/tickets";

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            HttpClient client = new HttpClient();
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            Console.WriteLine("Get Ticket...");
            HttpResponseMessage response = client.PostAsync(url, content).Result;
            if (response.IsSuccessStatusCode)
            {
                string res = response.Content.ReadAsStringAsync().Result;
                var responseData = JsonConvert.DeserializeObject(response.Content.ReadAsStringAsync().Result);

                Console.Write(responseData);
                XmlDocument xd = new XmlDocument();
                xd = (XmlDocument)Newtonsoft.Json.JsonConvert.DeserializeXmlNode(res);
                DataSet dsJson = new DataSet();
                dsJson.ReadXml(new XmlNodeReader(xd));
                if (dsJson != null)
                {
                    if (dsJson.Tables.Contains("entry"))
                    {
                        foreach (DataRow row in dsJson.Tables["entry"].Rows)
                        {
                            strTicket = Convert.ToString(row["id"]);
                            Console.WriteLine("Ticket:" + strTicket);
                        }
                    }
                }

            }

            // Upload package  
            string strEventCallback = string.Empty;
            string result = null;
            string strFileName = "";
            // Upload package  
            using (var clients = new HttpClient())
            {
                
                try
                {
                    using (var formData = new MultipartFormDataContent())
                    {
                        foreach (HttpPostedFileBase fileItem in file)
                        {
                            if (file != null)
                            {
                                strFileName=fileItem.FileName;
                                formData.Add(new StreamContent(fileItem.InputStream), "filedata", fileItem.FileName);
                                break;
                            }
                        }
                      
                        //formData.Add(new StringContent("mysiteid"), "siteid");
                        //formData.Add(new StringContent("mycontainerid"), "containerid");
                        //formData.Add(new StringContent("/"), "MyFolder");
                        formData.Add(new StringContent("cm:description"), "File Description");
                        //formData.Add(new StringContent("cm:content"), "Zip Package");
                        //formData.Add(new StringContent("cm:title"), "Books Final Package Automatic Ingestion");
                        formData.Add(new StringContent("true"), "overwrite");

                        //clients.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArrayCredential));

                        //Ticket based authentication
                        var byteArrayCre = Encoding.ASCII.GetBytes(strTicket);
                        clients.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArrayCre));
                        clients.Timeout = new TimeSpan(1, 5, 0);
             
                        var responses = clients.PostAsync(strURL, formData).Result;  
       
                        if (responses.Content != null)
                        {
                            result = responses.Content.ReadAsStringAsync().Result;
                            strEventCallback = result;
                            
                        }

                        if (responses.IsSuccessStatusCode)
                        {
                            if (string.IsNullOrWhiteSpace(result))
                                result = "Failed";
                            else
                            {
                                result = "Success";

                                string StrQry = "IF NOT EXISTS (Select JBM_AutoID, AutoArtID ArticleID from BK_ExternalEventAccess where AutoArtID='" + strAutoArtID + "' and ArticleID='" + strProjectID + "' and ExternalEventID='EV056' and FileName='" + strFileName + "') BEGIN INSERT INTO [BK_ExternalEventAccess] (JBM_AutoID, AutoArtID, ExternalEventID, ExternalEventStatus, ExternalEventActDate, ExternalEventTriggerd, StatusDescription, ArticleID,MaxTry,Status,Stage,FileName) VALUES ('" + strJBMAID + "', '" + strAutoArtID + "','EV056', 'Success' , GETDATE(), GETDATE(), '" + strEventCallback.Replace("'", "''") + "','" + strProjectID + "',1,1,'FML','" + strFileName + "') END ELSE BEGIN UPDATE [BK_ExternalEventAccess] SET ExternalEventStatus = 'Success',StatusDescription='" + strEventCallback.Replace("'", "''") + "', Status=1,ExternalEventTriggerd=GETDATE() WHERE AutoArtID='" + strAutoArtID + "' and ArticleID='" + projectid + "' and ExternalEventID='EV056' and FileName='" + strFileName + "' END";
                                string strResult = DBProc.GetResultasString(StrQry, "dbConnSmartCH-LIVE");
                                string strResult1 = DBProc.GetResultasString("INSERT INTO JBM_ProdAccess (EmpAutoID,CustAcc,AccPage,AccTime,Process,AutoArtID,JBM_AutoID,Descript) VALUES ('" + Session["EmpAutoId"].ToString()  + "','BK','Package Upload Page',GETDATE(),'Emerald API Upload','" + strAutoArtID  + "','" + strJBMAID  + "','Package uploaded thru API upload page. Package Name: " + strFileName + "')", "dbConnSmartCH-LIVE");
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(result))
                                result = "Failed";
                        }
 
                    }
                }
                catch (Exception ex)
                {
                    result = "Failed";
                }
                finally
                {
                }
            }

            return Json(new { dataUp = result, data1 = strEventCallback }, JsonRequestBehavior.AllowGet);

           // return View();
        }
     }
}