using SmartTrack.Helper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using System.Xml;

namespace SmartTrack.Controllers
{
    public class HelpDeskController : Controller
    {
        DataProc DBProc = new DataProc();
        clsCollection clsCollec = new clsCollection();
        DataTable dt;
        DataSet ds;
        clsINIst stINI = new clsINIst();
        string strINIPath = "";

        public HelpDeskController()
        {
            try
            {

                string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                XmlDocument xml = new XmlDocument();
                xml.Load(strPath);
                XmlNode xlNode = xml.SelectSingleNode("//config/AppINI-" + GlobalVariables.strEnvironment);
                strINIPath = xlNode.InnerText;

            }
            catch (Exception ex)
            {
                strINIPath = "";
            }
        }
           


    // GET: HelpDesk
    public ActionResult Index()
        {
            return View();
        }

        [SessionExpire]
        public ActionResult Tickets() 
        {
            DataProc DBProc = new DataProc();

            List<SelectListItem> items = new List<SelectListItem>();
            string strQuery = "Select DeptCode, DeptName, DeptSN from JBM_Departmentmaster";
            ds = new DataSet();
            
            ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow myRow in ds.Tables[0].Rows)
                {
                    items.Add(new SelectListItem
                    {
                        Text = myRow["DeptName"].ToString(),
                        Value = myRow["DeptCode"].ToString()
                    });
                }
            }

            //List<SelectListItem> Empitems = new List<SelectListItem>();
            //ds = new DataSet();
            //ds = DBProc.GetResultasDataSet("Select EmpLogin, EmpName from JBM_Employeemaster", Session["sConnSiteDB"].ToString());
            //if (ds.Tables[0].Rows.Count > 0)
            //{
            //    foreach (DataRow myRow in ds.Tables[0].Rows)
            //    {
            //        Empitems.Add(new SelectListItem
            //        {
            //            Text = myRow["EmpName"].ToString(),
            //            Value = myRow["EmpLogin"].ToString()
            //        });
            //    }
            //}

            ViewBag.deptList = items;
           // ViewBag.empList = Empitems;
            return View();
            
        }

        public ActionResult LoadEmployeeList(string id)
        {
            try
            {
                ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select EmpLogin, EmpName from JBM_Employeemaster  WHERE DeptCode in  ('" + id.Replace(",", "','") + "')", Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString(), a[1].ToString()
                 };

                return Json(new { dataEmp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { dataEmp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
          
        }
        
        //CreateNewTicket
        [HttpPost]
        
        public ActionResult Tickets(List<HttpPostedFileBase> fileUpload, string vEmpList, string vRT,string vPriority, string vSubject, string vQueries)  //string emp, List<HttpPostedFileBase> fileUpload 
        {
            try
            {

                vQueries = System.Uri.UnescapeDataString(vQueries);

                int result;
                result = DBProc.execNonQuery("INSERT INTO [JBM_HelpDeskTicket]  (RequestType,Subject,Description,RequestBy,RequestDate,Priority) Values ('" + vRT + "','" + vSubject + "','" + vQueries + "','" + Session["UserID"].ToString() + "', GETDATE(), '" + vPriority + "')", Session["sConnSiteDB"].ToString());

                if (result > 0)
                {
                    // Get the 
                    string strTicketID;
                    strTicketID = "SR" + DBProc.GetResultasString("SELECT IDENT_CURRENT('JBM_HelpDeskTicket')", Session["sConnSiteDB"].ToString());

                    string[] splitEmp = vEmpList.Split(new Char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string empID in splitEmp)
                    {
                        result = DBProc.execNonQuery("INSERT INTO [JBM_HelpDeskAlerts]  (TicketID,ReceivedIDs,MarkAsRead) Values ('" + strTicketID + "','" + empID + "','1')", Session["sConnSiteDB"].ToString());
                    }

                    /// Uploaded Files save and add the details in db
                    if (fileUpload != null)
                    {
                        string folderPath = Server.MapPath("~/UploadedFiles/Tickets/" + strTicketID);

                        if (!Directory.Exists(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                        }


                        //Uploaded file save server location with ticket id
                        foreach (HttpPostedFileBase file in fileUpload)
                        {
                            if (file != null)
                            {
                                string SavefilePath = Path.Combine(Server.MapPath("~/UploadedFiles/Tickets/" + strTicketID), file.FileName);
                                System.IO.File.WriteAllBytes(SavefilePath, clsCollec.ReadData(file.InputStream));
                            }

                        }
                        string path = "/UploadedFiles/Tickets/" + strTicketID;
                        string UpdateAttachmentPath = "UPDATE JBM_HelpDeskTicket SET ScreenShots='" + path + "' where TicketID='" + strTicketID + "'";
                        DBProc.UpdateRecord(UpdateAttachmentPath, Session["sConnSiteDB"].ToString());
                    }
                    // Returns message that successfully uploaded  
                    return Json("Ticket raised successfully!");
                }
                else
                {
                    return Json("Error in Ticket creation!");
                }

               
            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }

           
        }

        [SessionExpire]
        public ActionResult JiraTicket()
        {
            try
            {
                string strQuery = "select EmpMailId from JBM_EmployeeMaster where EmpAutoID='" + Session["EmpAutoId"].ToString() + "' select Description from JBM_EmployeeConfig where ProcessID='HID0001' and EmpAutoID='" + Session["EmpAutoId"].ToString() + "' ";

                DataSet ds = new DataSet();

                ds = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
                string EmpEmailID = "";
                string SupportEmailID = "";

                if (ds.Tables[0].Rows.Count > 0)
                {
                    EmpEmailID = ds.Tables[0].Rows[0][0].ToString();
                }

                if (ds.Tables[1].Rows.Count > 0)
                {
                    SupportEmailID = ds.Tables[1].Rows[0][0].ToString();
                }
                ViewBag.EmpEmailID = EmpEmailID;
                ViewBag.SupportEmailID = SupportEmailID;
            }
            catch (Exception)
            {

            }
            return View();
        }

        [ValidateInput(false)]
        public string JiraTicketSend(string from, string to,string cc, string subject, string query, string fileUploadPath, string mailContent, List<HttpPostedFileBase> attachment)
        {
            try
            {
                string[] arrAttachmentPath = new string[0];
                string errMsg = "";

                if (attachment != null)
                {
                    string strPath = Path.Combine(fileUploadPath, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    if (!Directory.Exists(strPath))
                    {
                        Directory.CreateDirectory(strPath);
                    }

                    int index = 0;
                    arrAttachmentPath = new string[attachment.Count];
                    foreach (HttpPostedFileBase attachmentFile in attachment)
                    {
                        arrAttachmentPath[index] = Path.Combine(strPath, attachmentFile.FileName);
                        attachmentFile.SaveAs(arrAttachmentPath[index]);
                        index++;
                    }
                }

                try
                {
                    MailMessage mailMsg = new MailMessage();
                    SmtpClient smtp = new SmtpClient();
                    mailMsg.From = new MailAddress(from);
                    mailMsg.To.Add(new MailAddress(to));

                    if (cc != "")
                    {
                        if (cc.IndexOf(";") > 0)
                        {
                            foreach (string ccID in cc.Split(';'))
                            {
                                mailMsg.CC.Add(new MailAddress(ccID));
                            }
                        }
                        else
                        {
                            mailMsg.CC.Add(new MailAddress(cc));
                        }
                    }

                    mailMsg.Subject = subject;
                    mailMsg.IsBodyHtml = true;
                    mailMsg.Body = mailContent;
                    mailMsg.Priority = MailPriority.High;

                    smtp.Host= clsINI.GetProfileString("MailServer", "IP", strINIPath);
                    string UseDefaultCredentials = clsINI.GetProfileString("MailServer", "DefaultCredentials", strINIPath);
                    smtp.UseDefaultCredentials = Convert.ToBoolean(UseDefaultCredentials);

                    if (smtp.UseDefaultCredentials==false)
                    {
                        string strUserName = clsINI.GetProfileString("MailServer", "UserName", strINIPath);
                        string strPassword = clsINI.GetProfileString("MailServer", "Password", strINIPath);

                        smtp.Credentials = new NetworkCredential(strUserName, strPassword);
                    }

                    //smtp.Host = "10.22.11.32";
                    //smtp.UseDefaultCredentials = false;

                    //string strUserName = "misblr";
                    //string strPassword = "pass@123";

                    //smtp.Credentials = new NetworkCredential(strUserName, strPassword);
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                    foreach (string PathToAttachment in arrAttachmentPath)
                    {
                        mailMsg.Attachments.Add(new Attachment(PathToAttachment));
                    }

                    smtp.Send(mailMsg);
                }catch(Exception ex)
                {
                    errMsg = "Error : " + ex.Message;
                }

                //ReferenceLibrary.clsMail.SendMail(from, to, "", "", from, subject, mailContent, ref errMsg, true, ReferenceLibrary.EmailPriority.High, arrAttachmentPath, "");

                if (errMsg == "")
                {
                    errMsg = "Success";
                }

                return errMsg;
            }
            catch (Exception ex)
            {
                return "Error : " + ex.Message;
            }
        }


    }
}