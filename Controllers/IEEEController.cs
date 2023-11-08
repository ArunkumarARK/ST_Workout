using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http.Results;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Newtonsoft.Json;
using SmartTrack.Helper;
using SmartTrack.Models;

namespace SmartTrack.Controllers
{
   
    public class IEEEController : Controller
    {
        // GET: IEEE
        [SessionExpire]
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        [AllowAnonymous]
        public ActionResult getIEEEMemberStatus(string strMemberID)
        {
            try
            {
                string strAccessToken = string.Empty;

                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("MemberID", "Rajachitra.S@kwglobal.com");
                    var content = new FormUrlEncodedContent(new[]
                    {
                    new KeyValuePair<string, string>("grant_type",  "client_credentials"),
                    new KeyValuePair<string, string>("client_id", "1e022f52-a2ed-4963-8dc8-c0579b43ad6d" ),
                    new KeyValuePair<string, string>("client_secret", "82e54b86-7c2d-4702-8efc-362044417c50"),
                    new KeyValuePair<string, string>("scope", "GetMemberStatus")
                });

                    var result = client.PostAsync("https://services13.ieee.org/RST/api/oauth/token", content).Result;
                    if (result.IsSuccessStatusCode)
                    {
                        string res = result.Content.ReadAsStringAsync().Result;
                        XNode node = JsonConvert.DeserializeXNode(res, "Root");
                        XmlNodeList objNode;
                        XmlDocument objXml = new XmlDocument();
                        objXml.LoadXml(node.ToString());
                        if (objXml.InnerXml != "Nothing")
                        {
                            objNode = objXml.SelectNodes("//Root/access_token");
                            if (objNode.Count > 0)
                            {
                                strAccessToken = objNode.Item(0).InnerText.ToString();
                            }
                        }
                    }

                }



                if (strAccessToken != "")
                {
                    HttpClient clientIeee = new HttpClient();
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

                    clientIeee.DefaultRequestHeaders.Add("authorization", "Bearer " + strAccessToken);
                    // string Jsonbody = "{\"MemberID\":\"haibochen@sjtu.edu.cn\"}";
                    string Jsonbody = "{\"MemberID\":\"" + strMemberID + "\"}";
                    var content = new StringContent(Jsonbody, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = clientIeee.PostAsync("https://services13.ieee.org/RST/Customer/getstatus", content).Result;
                    if (response.IsSuccessStatusCode)
                    {

                        string getStatus = response.Content.ReadAsStringAsync().Result;
                        XNode node = JsonConvert.DeserializeXNode(getStatus, "Root");
                        XmlNodeList objNode;
                        XmlDocument objXml = new XmlDocument();

                        string strMember = ""; string strGrade = ""; string strSocietyList = ""; ; string strLname = ""; string strFname = "";

                        objXml.LoadXml(node.ToString());
                        if (objXml.InnerXml != "Nothing")
                        {
                            objNode = objXml.SelectNodes("//Root/MemberStatus");
                            if (objNode.Count > 0)
                            {
                                strMember = objNode.Item(0).InnerText.ToString();
                            }
                            else
                            {
                                string strMessage = string.Empty;
                                objNode = objXml.SelectNodes("//Root/reasons/message");
                                if (objNode.Count > 0)
                                {
                                    return Json(new { dataNote = objNode.Item(0).InnerText.ToString() }, JsonRequestBehavior.AllowGet);
                                }
                            }

                            objNode = objXml.SelectNodes("//Root/Grade");
                            if (objNode.Count > 0)
                            {
                                strGrade = objNode.Item(0).InnerText.ToString();
                            }
                            objNode = objXml.SelectNodes("//Root/SocietyList");
                            if (objNode.Count > 0)
                            {
                                strSocietyList = objNode.Item(0).InnerText.ToString();
                            }
                            objNode = objXml.SelectNodes("//Root/LastName");
                            if (objNode.Count > 0)
                            {
                                strLname = objNode.Item(0).InnerText.ToString();
                            }
                            objNode = objXml.SelectNodes("//Root/FirstName");
                            if (objNode.Count > 0)
                            {
                                strFname = objNode.Item(0).InnerText.ToString();
                            }

                            var jsonResult = Json(new { dataNote = "Success", data1 = strMember, data2 = strSocietyList, data3 = strLname, data4 = strFname, data5 = strGrade, }, JsonRequestBehavior.AllowGet);

                            jsonResult.MaxJsonLength = int.MaxValue;

                            return jsonResult;
                        }
                    }
                }
                else
                {
                    return Json(new { dataNote = "Getting access token failed." }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { dataNote = ex.Message }, JsonRequestBehavior.AllowGet);
            }

            return Json(new { dataNote = "Success" }, JsonRequestBehavior.AllowGet);

        }

    }
}