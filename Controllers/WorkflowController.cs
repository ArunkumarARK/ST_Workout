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

namespace SmartTrack.Controllers
{
    public class WorkflowController : Controller
    {
        clsCollection clsCollec = new clsCollection();
        clsINIst stINI = new clsINIst();
        DataProc DBProc = new DataProc(); // Data store/retrive DB
        SmartTrack.DataSecurity objDS = new SmartTrack.DataSecurity();  // For Encrypt/Decrypt
        Generic gen = new Generic();
        // GET: Workflow
        [SessionExpire]
        public ActionResult ProjectWF()
        {
            

            try
            {
                Dictionary<string, string> wfCollection = new Dictionary<string, string>();
                string strSqlFinal = string.Empty;
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet("Select Stages from [dbo].[JBM_AccountTypeDesc] where CustAccess='BK'", Session["sConnSiteDB"].ToString());
                if (ds.Tables[0].Rows.Count > 0)
                {
                    
                    string strMenuItem = ds.Tables[0].Rows[0]["Stages"].ToString();
                    string[] strComboValues = strMenuItem.Split('|');
                    strSqlFinal = "Select ";

                    for (int I = 1; I <= strComboValues.Length - 2; I++)
                    {
                        string[] strTmp = strComboValues[I].ToString().Split('-');
                        if (strTmp[0] != null)
                        {
                            strSqlFinal += strTmp[1] + "_wf as [" + strTmp[0] + "], ";
                            if (wfCollection.ContainsKey(strTmp[0]) == false)
                            {
                                wfCollection.Add(strTmp[0], strTmp[1]);
                            }
                           
                        }
                           

                    }

                    int cpos = strSqlFinal.Trim().ToString().LastIndexOf(",");
                    strSqlFinal = strSqlFinal.Trim().ToString().Substring(0, cpos) + " from JBM_Info where JBM_AutoId='" + Session["sJBMAutoID"].ToString() + "' ";
                }

                string strConcanteWFCode = string.Empty;
                string strConcanteWFCodeDesc = string.Empty;
                DataSet dsWFCode = new DataSet();
                string strQuery = "";
                if (GlobalVariables.strSiteCode == "ND")
                {
                    strQuery = "Select WFName from JBM_WFCode  where ND_Workflow = '1' order by cast(substring(WFName, 2, 5) as int) asc";
                }
                else
                {
                    strQuery = "Select WFName from JBM_WFCode order by cast(substring(WFName, 2, 5) as int) asc";
                }

                dsWFCode = DBProc.GetResultasDataSet(strQuery, Session["sConnSiteDB"].ToString());
                if (dsWFCode.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsWFCode.Tables[0].Rows.Count; intCount++)
                    {
                        strConcanteWFCode += dsWFCode.Tables[0].Rows[intCount]["WFName"].ToString() + "+";

                        string strWFDescNames = "";
                        DataSet dsWName = new DataSet();
                        dsWName = DBProc.GetResultasDataSet("Select dbo.Split_Function_WF('" + dsWFCode.Tables[0].Rows[intCount]["WFName"].ToString() + "') as Wf", Session["sConnSiteDB"].ToString());
                        if (dsWName.Tables[0].Rows.Count > 0)
                        {
                            strWFDescNames = dsWName.Tables[0].Rows[0][0].ToString();
                            strWFDescNames = strWFDescNames.Replace(",", @"  >  ");
                        }
                        strConcanteWFCodeDesc += strWFDescNames + "|";

                    }

                    int cpos = strConcanteWFCode.Trim().ToString().LastIndexOf("+");
                    strConcanteWFCode = strConcanteWFCode.Trim().ToString().Substring(0, cpos);

                    int wpos = strConcanteWFCodeDesc.Trim().ToString().LastIndexOf("|");
                    strConcanteWFCodeDesc = strConcanteWFCodeDesc.Trim().ToString().Substring(0, wpos);
                }

                DataSet dsNew = new DataSet();
                dsWFCode = new DataSet();
                dsWFCode = DBProc.GetResultasDataSet(strSqlFinal, Session["sConnSiteDB"].ToString());
                if (dsWFCode.Tables[0].Rows.Count > 0)
                {
                
                    DataTable dt = new DataTable("MyTable");
                    dt.Columns.Add(new DataColumn("DisplayName", typeof(string)));
                    dt.Columns.Add(new DataColumn("SelectedItem", typeof(string)));
                    dt.Columns.Add(new DataColumn("Dropdownlist", typeof(string)));
                    dt.Columns.Add(new DataColumn("ShortName", typeof(string)));
                    dt.Columns.Add(new DataColumn("ViewColumn", typeof(string)));
                    dt.Columns.Add(new DataColumn("Tooltip", typeof(string)));

                    DataRow dr = dt.NewRow();
                    string strDisplayName = "";
                    string strSelectedItem = "";
                    string strDropdownlist = "";
                    string strShortName = "";
                    string strWFDescList = "";

                    for (int intCount = 0; intCount < dsWFCode.Tables[0].Rows.Count; intCount++)
                    {

                        int vCol = 1;
                        foreach (DataColumn col in dsWFCode.Tables[0].Columns)
                        {
                            if (vCol == 4)
                            {
                                vCol = 1;
                            }

                            strDisplayName = col.ColumnName.ToString();
                            strSelectedItem = dsWFCode.Tables[0].Rows[intCount][col.ColumnName].ToString();
                            strDropdownlist = strConcanteWFCode;
                            strWFDescList = strConcanteWFCodeDesc;

                            if (wfCollection.ContainsKey(strDisplayName))
                            {
                                strShortName = wfCollection[strDisplayName];
                            }

                            dr = dt.NewRow();
                            dr["DisplayName"] = strDisplayName;
                            dr["SelectedItem"] = strSelectedItem;
                            dr["Dropdownlist"] = strDropdownlist;
                            dr["ShortName"] = strShortName;
                            dr["ViewColumn"] = vCol.ToString();
                            dr["Tooltip"] = strWFDescList.ToString();

                            dt.Rows.Add(dr);

                            vCol = vCol + 1;

                        }

                    }

                   

                    dsNew.Tables.Add(dt);
                    ViewBag.dt = dt;

                }

               


            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }
        
        public DataTable SetRecord_All(string strFieldVal)
        {
            DataTable strOutRecordCol = new DataTable();
            strOutRecordCol = Proc_Set_FieldVal(strFieldVal);
            return strOutRecordCol;
        }
        public DataTable Proc_Set_FieldVal(string strTemp)
        {
            DataTable strOut = new DataTable();
            //Dim FieldVal As Init_ColVal
            //FieldVal = Nothing

            string[] strIn = strTemp.Split('|'); //Split(strTemp, "|", -1, CompareMethod.Binary)

            try
            {
                DataRow dtrow = strOut.NewRow();// DataRow();
                strOut.Rows.Add(dtrow);
                for (int i = 0; i < strIn.Length; i++)
                {
                    //    FieldVal = Nothing  
                    string[] strInvalues = strIn[i].Split(',');
                    //    FieldVal.gColName = Mid(strIn(i), 1, InStr(1, strIn(i), ",", CompareMethod.Binary) - 1) 
                    //    Dim strVal As String = Mid(strIn(i), InStr(1, strIn(i), ",", CompareMethod.Binary) + 1)   
                    if (strInvalues.Length > 0)
                    {
                        string colval = strInvalues[0].ToString();
                        string dataval = strInvalues[1].ToString();
                        strOut.Columns.Add(colval, typeof(string));
                        DataRow dr = strOut.Rows[0];
                        dr[colval] = dataval;

                    }

                }
            }
            catch (Exception ex)
            {
                // Result = "Error";
            }
            finally
            {
                //FieldVal = Nothing;
            }

            return strOut;
        }
        [SessionExpire]
        public JsonResult UpdateWorkflowDetails(string wfCodes)
        {
            try
            {
                SqlConnection con = new SqlConnection();
                con = DBProc.getConnection(Session["sConnSiteDB"].ToString());
                con.Open();
                DataTable clsCustSave = new DataTable();
                clsCustSave = new DataTable();
                string strQuery = string.Empty;
                string strFieldVal = string.Empty;
                string strAutoJID = Session["sJBMAutoID"].ToString();
                string strStageValue = "";

                // Update JBM Info'
                strQuery = "Update JBM_Info set ";

                List<string> wfSaveList = JsonConvert.DeserializeObject<List<string>>(wfCodes);
                if (wfSaveList.Count > 0)
                {
                    for (int j = 0; j < wfSaveList.Count; j++)
                    {
                        string strShortName = wfSaveList[j].Split('|')[0];
                        string strWFCode = wfSaveList[j].Split('|')[1];

                        string strHeaderVal = strShortName + "_wf";
                        string strComboVal = strWFCode;

                        // Suresh Workflow modified based Stage
                        // string strlWorkflowMsg = lblWorkflowMsg.Text.Trim;
                        string[] strsplitwrkflow = wfSaveList.ToArray();// strlWorkflowMsg.Split(',');

                        for (int i = 0; i <= strsplitwrkflow.Length - 1; i++)
                        {
                            if (strsplitwrkflow[i].Trim().Contains(strShortName.ToString().Trim()))
                            {
                                string[] myDelims = new string[] { "|" };
                                string[] StageField = strsplitwrkflow[i].Split(myDelims, StringSplitOptions.None);
                                string StageName = StageField[0].Trim().ToString();
                                string StageValue = StageField[1].Trim().ToString();
                                if (StageValue.Trim() == "")
                                    StageValue = null;
                                if (strComboVal != StageValue && StageName == strShortName)
                                {
                            string strstageType = "";

                            DataTable dt1 = DBProc.GetResultasDataTbl("Select StageTypeid,FP_dependancy from JBM_StageDescription where stageshortname='" + StageName + "'", Session["sConnSiteDB"].ToString());

                            if (dt1.Rows.Count != 0)
                            {
                                strstageType = dt1.Rows[0]["StageTypeid"].ToString();

                                if (dt1.Rows[0]["FP_dependancy"].ToString() == "" | dt1.Rows[0]["FP_dependancy"].ToString() == "1")
                                    strstageType += "|1";
                                else
                                    strstageType += "|0";
                            }
                            string StageTypeID = strstageType;
                                    string strStagequery = "";
                                    // ArtStage based Workflow
                                    DataTable DtWFCode = DBProc.GetResultasDataTbl("Select WFCode,CeParallel,XMLParallel from JBM_WFCode where WFName='" + strComboVal + "'", Session["sConnSiteDB"].ToString());
                                    string strArtStage = "";
                                    if (DtWFCode.Rows.Count !=0 )
                                    {
                                        string strCode = DtWFCode.Rows[0][0].ToString();
                                        string[] strInput = strCode.Split('|');
                                        string strOrder = strInput[0];
                                        string[] strIn = strOrder.Split(',');
                                        string strPart1 = "";
                                        string strPart2 = "";
                                        if (strIn[0].Length == 1)
                                            strPart1 = "S100";
                                        if (strIn[0].Length == 2)
                                            strPart1 = "S10";
                                        if (strIn[0].Length == 1)
                                            strPart2 = "S100";
                                        if (strIn[1].Length == 2)
                                            strPart2 = "S10";
                                        strArtStage = ",ArtStageTypeID='" + strPart1 + strIn[0] + "_S1010_" + strPart2 + strIn[1] + "_" + Session["EmpIdLogin"] + "_" + System.DateTime.Now.ToString() + "'";
                                    }


                                    if (StageTypeID.Substring(0, 1).ToString() == "1")
                                        strStagequery = "update " + Session["sCustAcc"].ToString() + "_stageinfo  set Rev_wf=" + (strComboVal == ""? "NULL": "'" + strComboVal + "'") + strArtStage + " from " + Session["sCustAcc"].ToString() + "_stageinfo  s, " + Session["sCustAcc"].ToString() + "_chapterinfo  a where a.AutoArtID=s.AutoArtID and a.JBM_AutoID='" + strAutoJID + "' and s.RevFinStage like '" + StageName + "%' ";
                                    else if (StageTypeID.Substring(0, 1).ToString() == "2")
                                        strStagequery = "update " + Session["sCustAcc"].ToString() + "_IssueInfo set Rev_wf=" + (strComboVal == ""? "NULL": "'" + strComboVal + "'") + strArtStage + " from " + Session["sCustAcc"].ToString() + "_IssueInfo s," + Session["sCustAcc"].ToString() + "_chapterinfo a where a.JBM_AutoID=s.JBM_AutoID and isnull(a.iss,'9999999')=isnull(s.iss,'9999999') and a.JBM_AutoID='" + strAutoJID + "' and  s.RevFinStage like '" + StageName + "%' ";
                           
                                    SqlCommand cmdstrStagequery = new SqlCommand(strStagequery, con);
                                    cmdstrStagequery.ExecuteNonQuery();
                                    strStagequery = "";
                                    break;
                                }
                            }
                        }
                        // Workflow modified based Stage

                        strFieldVal += strHeaderVal + "," + strComboVal + "|";
                        clsCustSave = SetRecord_All(strFieldVal);
                        string stritmValue = "'"+clsCustSave.Rows[0][strHeaderVal].ToString().Trim()+"'";
                        strQuery += strHeaderVal + "=" + stritmValue + ", ";
                        strStageValue += strHeaderVal + "=" + stritmValue + ", ";
                        DataTable dtsCustGroup = DBProc.GetResultasDataTbl("select CustGroup from jbm_customermaster where CustType='" + Session["sCustAcc"].ToString() + "' and CustSN='" + Session["sCustSN"].ToString() + "'", Session["sConnSiteDB"].ToString());

                        if (strHeaderVal.ToLower() == "sample_wf" & dtsCustGroup.Rows[0]["CustGroup"].ToString()== "CG001")
                        {
                            strFieldVal += "srev_wf" + "," + strComboVal + "|";
                            strQuery += "srev_wf" + "=" + stritmValue + ", ";
                            strStageValue += "srev_wf" + "=" + stritmValue + ", ";
                        }

                    }
                    int cpos = strQuery.LastIndexOf(",");
                    strQuery = strQuery.Substring(0, cpos);
                    strQuery += " where JBM_AutoID='" + strAutoJID + "'";
                   
                    SqlCommand cmdstrQuery = new SqlCommand(strQuery, con);
                    cmdstrQuery.ExecuteNonQuery();
                   
                   

                    // update JBM StageInfo'

                    strQuery = "select s.autoartid, s.revfinstage,s.rev_wf from " + Session["sCustAcc"].ToString() + "_stageinfo s join " + Session["sCustAcc"].ToString() + "_chapterinfo c on s.autoartid=c.autoartid where c.JBM_AutoID ='" + strAutoJID + "' and s.DispatchDate is null";
                    strStageValue = strStageValue.Replace("_wf", "");
                    string[] stageArr = strStageValue.Split(',');
                    string strUpdateQuery = "";
                    string wf = "";
                   
                    DataTable dt = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
                    if (dt.Rows.Count != 0)
                    {
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {
                            for (int j = 0; j <= stageArr.Length - 1; j++)
                            {
                                if (stageArr[j].Contains(dt.Rows[i][1].ToString()))
                                {
                                    string[] stageArrs = stageArr[j].Split('=');

                                    if (stageArrs[1] == "" | stageArrs[1] == "Null")
                                        wf = "Null";
                                    else
                                    {
                                        // wf = Strings.Mid(stageArr[j], Strings.InStr(Strings.InStr(stageArr[j], "=", Constants.vbTextCompare), stageArr[j], "W", Microsoft.VisualBasic.CompareMethod.Binary) - 1);
                                        string[] wfarray = stageArr[j].Split('=');
                                        wf = wfarray[1].ToString();
                                    }
                                    strUpdateQuery = "update " + Session["sCustAcc"].ToString() + "_stageinfo set Rev_wf=" + wf + " where AutoArtID='" + dt.Rows[i][0].ToString() + "' and RevFinStage like '%" + dt.Rows[i][1].ToString() + "%' and DispatchDate is null";
                                    
                                    SqlCommand cmdstrUpdateQuery = new SqlCommand(strUpdateQuery, con);
                                    cmdstrUpdateQuery.ExecuteNonQuery();
                                   
                                   // RecordManager.UpdateRecord(strUpdateQuery);

                                    break;
                                }
                            }
                        }
                    }

                    

                    // Update Issue Info"
                    for (int j = 0; j <= stageArr.Length - 1; j++)
                    {
                        if (stageArr[j].Contains("Fin"))
                        {
                            string[] wfarray = stageArr[j].Split('=');
                            wf = wfarray[1].ToString();
                           // wf = Strings.Mid(stageArr[j], Strings.InStr(Strings.InStr(stageArr[j], "=", Constants.vbTextCompare), stageArr[j], "W", Microsoft.VisualBasic.CompareMethod.Binary) - 1);
                            strUpdateQuery = "update " + Session["sCustAcc"].ToString() + "_IssueInfo set Rev_wf=" + wf + " where JBM_AutoID='" + strAutoJID + "' and DispatchDate is null";
                            SqlCommand cmdstrUpdateQuery = new SqlCommand(strUpdateQuery, con);
                            cmdstrUpdateQuery.ExecuteNonQuery();
                           // RecordManager.UpdateRecord(strUpdateQuery);

                            break;
                        }
                    }
                    con.Close();

                }
                else
                {
                    return Json(new { dataWf = "Failed" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataWf = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataWf = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult getWrokflowDetails()
        {

            try
            {
                string strCWP = "";
                string strSelectWFTypes = "";
                if (strCWP.Contains("CWP"))
                {
                    strSelectWFTypes = "Select WFCode ,WFName,CeParallel,CeIndex, XMLParallel,MyPetParallel, XMLIndex, EmpAutoID, cwp, ND_Workflow from JBM_WFCode where CWP is not null order by cast(substring(WFName, 2,5) as int)";
                }
                else {
                    strSelectWFTypes = "Select WFCode ,WFName,CeParallel,CeIndex, XMLParallel,MyPetParallel, XMLIndex, EmpAutoID, cwp, ND_Workflow from JBM_WFCode where ND_Workflow=1 and CWP is null order by cast(substring(WFName, 2,5) as int)";
                }
            
                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strSelectWFTypes, Session["sConnSiteDB"].ToString());

                if (ds.Tables[0].Rows.Count > 0)
                {

                    string wfTableList = @"<table id=tblwf style=""width:auto;"">";
                    //wfTableList += @"<tr class=""lockHeadDivgrid h3"" align=""center"" style=""background-color: #f5f5dc;""><td align=center colspan=2><font color=#FF8000><strong class=h3>Existing&nbsp;Workflows</strong></font></td></tr>";

                    for (int intCount = 0; intCount < ds.Tables[0].Rows.Count; intCount++)
                    {
                        string strWFName = ds.Tables[0].Rows[intCount]["WFName"].ToString();
                        string strWFDescNames = "";
                        DataSet dsWName = new DataSet();
                        dsWName = DBProc.GetResultasDataSet("Select dbo.Split_Function_WF('" + strWFName + "') as Wf", Session["sConnSiteDB"].ToString());
                        if (dsWName.Tables[0].Rows.Count > 0)
                        {
                            strWFDescNames = dsWName.Tables[0].Rows[0][0].ToString();
                            strWFDescNames = strWFDescNames.Replace(",", @"&nbsp;&nbsp;&nbsp;<i class=""fas fa-caret-right"" style=""color:red""></i>&nbsp;&nbsp;&nbsp;");
                        }

                        wfTableList += @"<tr><td style=""padding:3px;""><b><font color =#FF8000>" + strWFName + "</font></b>&nbsp;&nbsp;</td>";
                        wfTableList += @"<td style=""padding:3px;""><font color =>" + strWFDescNames + "</font></td></tr>";
                        wfTableList += @"<tr><td style=""padding:3px;"" colspan=2></td></tr>";
                    }

                    wfTableList += "</table>";
                    return Json(new { aaData = wfTableList }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { aaData = "NoRecord" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { aaData = ex.Message }, JsonRequestBehavior.AllowGet);
            }

           
        }


    }

}