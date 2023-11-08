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
    public class JobReportController : Controller
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
        // GET: JobHistory
        public ActionResult Index()
        {
            return View();
        }
        [SessionExpire]
        public ActionResult JobHistory()
        {
            try
            {
                ViewBag.PrjTabHead = "Job History";
                ViewBag.PageHead = "Job History";


                //Load Customer List
                List<SelectListItem> lstCust = new List<SelectListItem>();
                DataSet dsCust = new DataSet();
                dsCust = DBProc.GetResultasDataSet("select distinct jc.custname,jc.custid,jc.CustSN from JBM_CustomerMaster jc join jbm_info j on jc.custid=j.custid where j.jbm_disabled='0' and custtype = '" + Session["sCustAcc"].ToString() + "'  order by jc.custsn ", Session["sConnSiteDB"].ToString());
                if (dsCust.Tables[0].Rows.Count > 0)
                {
                    for (int intCount = 0; intCount < dsCust.Tables[0].Rows.Count; intCount++)
                    {
                        string strEmpAutoID = dsCust.Tables[0].Rows[intCount]["custid"].ToString();
                        string strEmpName = dsCust.Tables[0].Rows[intCount]["custname"].ToString();
                        string strCustSN = dsCust.Tables[0].Rows[intCount]["CustSN"].ToString();
                        lstCust.Add(new SelectListItem
                        {
                            Text = strCustSN, //strEmpName.ToString() + " (" + strCustSN + ")",
                            Value = strEmpAutoID.ToString()
                        });
                    }

                }

                ViewBag.Custlist = lstCust;

                DataSet dsartnew = new DataSet();
                DataSet dt = (DataSet)Session["dsart"];
                if (dt != null)
                {
                    dsartnew = dt;
                    Session["dsartexcel"] = dt;
                    Session["dsart"] = null;
                    Session["Lst"] = Session["List"];
                    Session["CstSelect"] = Session["CustSelect"];
                    Session["JDSelect"] = Session["JIDSelect"];
                    Session["IssSelect"] = Session["IssueSelect"];
                    Session["List"] = null;
                    return View(dsartnew);
                }
                else
                {
                    Session["dsartexcel"] = null;
                    Session["Lst"] = null;
                    Session["CstSelect"] = "0";
                    Session["JDSelect"] = "0";
                    Session["IssSelect"] = "0";
                    return View();
                }
            }
            catch (Exception ex)
            {
                return View();
            }

        }
        public void CreateExcelFile(DataSet data, string OutPutFileDirectory, string fileName)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(OutPutFileDirectory, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }
        }
        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)3U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            // FillId = 2,RED
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "93ebf9" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;

        }
        private void CreatePartsForExcel(SpreadsheetDocument document, DataSet data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData, data);

        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1, DataSet data)
        {
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet.Append(sheetDimension1);
            worksheet.Append(sheetViews1);
            worksheet.Append(sheetFormatProperties1);
            worksheet.Append(sheetData1);
            worksheet.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet;
            int count = 1;
            MergeTwoCells(worksheet, "A" + count, "M" + count);
            count++;
            for (int i = 0; i < data.Tables.Count; i++)
            {
                MergeTwoCells(worksheet, "A" + count, "M" + count);
                count = count + data.Tables[i].Rows.Count + 2;
            }
            SetColumnWidth(worksheet, 1U, 60);
        }
        static void SetColumnWidth(Worksheet worksheet, uint Index, DoubleValue dwidth)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns cs = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (cs != null)
            {
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Column> ic = cs.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(r => r.Min == Index).Where(r => r.Max == Index);
                if (ic.Count() > 0)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = ic.First();
                    c.Width = dwidth;
                }
                else
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                    cs.Append(c);
                }
            }
            else
            {
                cs = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                cs.Append(c);
                worksheet.InsertAfter(cs, worksheet.GetFirstChild<SheetFormatProperties>());
            }
        }
        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        private SheetData GenerateSheetdataForDetails(DataSet data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateMainHeaderRowForExcel("Main", ""));
            for (int i = 0; i < data.Tables.Count; i++)
            {
                string tablename = data.Tables[i].TableName;
                if (data.Tables[i].Rows.Count > 0)
                {
                    sheetData1.Append(CreateMainHeaderRowForExcel(tablename, data.Tables[i].Rows[0][0].ToString()));
                    sheetData1.Append(CreateHeaderRowForExcel(tablename));
                    foreach (DataRow taktTimemodel in data.Tables[i].Rows)
                    {
                        Row partsRows = GenerateRowForChildPartDetail(taktTimemodel, tablename);
                        sheetData1.Append(partsRows);
                    }
                }
            }

            return sheetData1;
        }
        private Row CreateHeaderRowForExcel(string tablename)
        {
            Row workRow = new Row();

            if (tablename == "Table")
            {
                workRow.Append(CreateCell("JobID"));
                workRow.Append(CreateCell("JID"));
                workRow.Append(CreateCell("AutoID"));
                workRow.Append(CreateCell("InternalID"));
                workRow.Append(CreateCell("ArticleID"));
                workRow.Append(CreateCell("MSID"));
                workRow.Append(CreateCell("Platform"));
                workRow.Append(CreateCell("No of Figures"));
                workRow.Append(CreateCell("No of Refs"));
                workRow.Append(CreateCell("MSP"));
                workRow.Append(CreateCell("TSP"));
                workRow.Append(CreateCell("No of Tables"));
                workRow.Append(CreateCell("Vender Pages"));
            }
            else if (tablename == "Table1")
            {
                workRow.Append(CreateCell("JobID"));
                workRow.Append(CreateCell("Stage"));
                workRow.Append(CreateCell("ArtStage"));
                workRow.Append(CreateCell("Current Stage"));
                workRow.Append(CreateCell("WF"));
                workRow.Append(CreateCell("Current Status"));
                workRow.Append(CreateCell("Received Date"));
                workRow.Append(CreateCell("Due Date"));
                workRow.Append(CreateCell("Dispatched Date"));
                workRow.Append(CreateCell("PDF GEN Start"));
                workRow.Append(CreateCell("PDF GEN End"));
            }
            else if (tablename == "Table2")
            {
                workRow.Append(CreateCell("InternalID"));
                workRow.Append(CreateCell("ChapterID"));
                workRow.Append(CreateCell("ISS"));
                workRow.Append(CreateCell("Platform"));
                workRow.Append(CreateCell("No of Figures"));
                workRow.Append(CreateCell("AutoArtID"));
                workRow.Append(CreateCell("JBM_Intrnl"));
                workRow.Append(CreateCell("CustSN"));
                workRow.Append(CreateCell("JBM_AutoID"));
            }
            else if (tablename == "Table3")
            {
                workRow.Append(CreateCell("JobID"));
                workRow.Append(CreateCell("ISS"));
                workRow.Append(CreateCell("Stage"));
                workRow.Append(CreateCell("Current Stage"));
                workRow.Append(CreateCell("WF"));
                workRow.Append(CreateCell("Received Date"));
                workRow.Append(CreateCell("Due Date"));
                workRow.Append(CreateCell("Dispatched Date"));
            }
            else if (tablename == "Table4")
            {
                workRow.Append(CreateCell("AutoArtID"));
                workRow.Append(CreateCell("Stage"));
                workRow.Append(CreateCell("Au Reminder"));
                workRow.Append(CreateCell("Ed Reminder"));
                workRow.Append(CreateCell("Pe Reminder"));
                workRow.Append(CreateCell("Pr Reminder"));
            }
            else if (tablename == "Table5")
            {
                workRow.Append(CreateCell("AutoArtID"));
                workRow.Append(CreateCell("Stage"));
                workRow.Append(CreateCell("Proof WF"));
                workRow.Append(CreateCell("Mail From"));
                workRow.Append(CreateCell("Author Email"));
                workRow.Append(CreateCell("Editor Email"));
                workRow.Append(CreateCell("PE Email"));
                workRow.Append(CreateCell("PR Email"));
                workRow.Append(CreateCell("Proof Type"));
            }
            else if (tablename == "Table6")
            {
                workRow.Append(CreateCell("AutoArtID"));
                workRow.Append(CreateCell("Stage"));
                workRow.Append(CreateCell("ToStop Licence Cancel"));
            }
            return workRow;
        }
        private Row CreateMainHeaderRowForExcel(string tablename, string ProjectID)
        {
            Row workRow = new Row();

            if (tablename == "Table")
            {
                workRow.Append(CreateCell("ARTICLE DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table1")
            {
                workRow.Append(CreateCell("STAGE DETAILS ", 1U, workRow.RowIndex));
            }
            else if (tablename == "Main")
            {
                workRow.Append(CreateCell("ARTICLE REPORT", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table2")
            {
                workRow.Append(CreateCell("CHAPTER DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table3")
            {
                workRow.Append(CreateCell("ISSUE DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table4")
            {
                workRow.Append(CreateCell("CANCEL REMINDER DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table5")
            {
                workRow.Append(CreateCell("PROOFING DETAILS", 1U, workRow.RowIndex));
            }
            else if (tablename == "Table6")
            {
                workRow.Append(CreateCell("CANCEL LICENCE DETAILS", 1U, workRow.RowIndex));
            }
            return workRow;
        }
        private Row GenerateRowForChildPartDetail(DataRow testmodel, string tablename)
        {
            Row tRow = new Row();

            if (tablename == "Table")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[5].ToString()));
                tRow.Append(CreateCell(testmodel[6].ToString()));
                tRow.Append(CreateCell(testmodel[10].ToString()));
                tRow.Append(CreateCell(testmodel[12].ToString()));
                tRow.Append(CreateCell(testmodel[13].ToString()));
                tRow.Append(CreateCell(testmodel[15].ToString()));
                tRow.Append(CreateCell(testmodel[16].ToString()));
                tRow.Append(CreateCell(testmodel[19].ToString()));
                tRow.Append(CreateCell(testmodel[20].ToString()));
            }
            else if (tablename == "Table1")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[20].ToString()));
                tRow.Append(CreateCell(testmodel[21].ToString()));
                tRow.Append(CreateCell(testmodel[22].ToString()));
                tRow.Append(CreateCell(testmodel[27].ToString()));
                tRow.Append(CreateCell(testmodel[28].ToString()));
            }
            else if (tablename == "Table2")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[5].ToString()));
                tRow.Append(CreateCell(testmodel[6].ToString()));
                tRow.Append(CreateCell(testmodel[7].ToString()));
                tRow.Append(CreateCell(testmodel[8].ToString()));
                tRow.Append(CreateCell(testmodel[10].ToString()));
            }
            else if (tablename == "Table3")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[6].ToString()));
                tRow.Append(CreateCell(testmodel[7].ToString()));
                tRow.Append(CreateCell(testmodel[8].ToString()));
            }
            else if (tablename == "Table4")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[5].ToString()));
            }
            else if (tablename == "Table5")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
                tRow.Append(CreateCell(testmodel[3].ToString()));
                tRow.Append(CreateCell(testmodel[4].ToString()));
                tRow.Append(CreateCell(testmodel[5].ToString()));
                tRow.Append(CreateCell(testmodel[6].ToString()));
                tRow.Append(CreateCell(testmodel[7].ToString()));
                tRow.Append(CreateCell(testmodel[8].ToString()));
            }
            else if (tablename == "Table6")
            {
                tRow.Append(CreateCell(testmodel[0].ToString()));
                tRow.Append(CreateCell(testmodel[1].ToString()));
                tRow.Append(CreateCell(testmodel[2].ToString()));
            }

            return tRow;
        }
        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 2U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private Cell CreateCell(string text, uint styleIndex, UInt32Value RowIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
        private static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name)
        {

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.

            string s1 = cell1Name + ":" + cell2Name;
            MergeCell mergeCell = new MergeCell() { Reference = s1 };
            mergeCells.Append(mergeCell);

            worksheet.Save();

        }
        private byte[] ConvertDataSetToByteArray(DataSet dataSet)
        {
            byte[] binaryDataResult = null;
            using (MemoryStream memStream = new MemoryStream())
            {
                BinaryFormatter brFormatter = new BinaryFormatter();
                dataSet.RemotingFormat = SerializationFormat.Binary;
                brFormatter.Serialize(memStream, dataSet);
                binaryDataResult = memStream.ToArray();
            }
            return binaryDataResult;
        }

        public FileResult ExcelReport()
        {
            DataSet ds = (DataSet)Session["dsartexcel"];
            var datetime = DateTime.Now.ToString("MMddyyyyHHmmss");
            string fileName = "Report_" + datetime + ".xlsx";
            string path = Server.MapPath("~/UploadedFiles/Report_" + datetime + ".xlsx");
            CreateExcelFile(ds, path, fileName);
            byte[] bytes;
            bytes = System.IO.File.ReadAllBytes(path);
            return File(bytes, "application/octet-stream", fileName);

        }
        public void GetHistoryLink(string strAutoart, string strMode, string txtID, string cboID, string rdSearchList, string cboIssues, string lblHis2, string isBookJnl)
        {
            string strCustAcc = Session["sCustAcc"].ToString();
            string CustDisplayId = "1";// clsInit.gStrDisplayCustIdPD;
            string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
            string strStageinfo = strCustAcc + Init_Tables.gTblStageInfo;
            string strProdStatus = strCustAcc + Init_Tables.gTblProdStatus;
            string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
            DataTable dtArticle = new DataTable();
            DataTable dtArticleDetails = new DataTable();
            string strCustGroup = Session["sCustGroup"].ToString();
            string strUserLoginName = Session["UserID"].ToString();
            string strCboVal = "";

            if (cboID != "")
            {
                strCboVal = "and a.jbm_autoid='" + cboID + "'";
            }
            else
            {
                string jobquery = "select jbm_autoid from " + strCustAcc + "_ArticleInfo where AutoArtID='" + txtID + "' or IntrnlID='" + txtID + "' or ChapterID='" + txtID + "' or Iss='" + txtID + "'";
                DataTable dtjob = new DataTable();
                dtjob = DBProc.GetResultasDataTbl(jobquery, Session["sConnSiteDB"].ToString());
                if (dtjob.Rows.Count > 0)
                {
                    cboID = dtjob.Rows[0][0].ToString();
                    strCboVal = "and a.jbm_autoid='" + cboID + "'";
                }
                else
                {
                    return;
                }
            }

            string strIss = "";
            string strArtSearchValue = "";
            string strSearchValue = "and (a.AutoArtID='" + txtID + "' or a.IntrnlID='" + txtID + "' or a.ChapterID='" + txtID + "' or a.DOI='" + txtID + "' ) "; //txtID.Text;
            strArtSearchValue = strSearchValue;
            if (rdSearchList == "2")
            {
                strIss = "Iss";
                if (cboIssues.Trim() != "[Select]")
                {
                    strSearchValue = "and a.iss='" + cboIssues + "'";
                    strArtSearchValue = "and (a.iss='" + cboIssues + "' or a.TempIss='" + cboIssues + "')";
                }
                //btnUpdateCID.Visible = false;

            }
            else if (CustDisplayId == "1" && cboID != "")
            {
                strIss = "IssBooks";

                if (txtID == "")
                {
                    strSearchValue = "";
                    strArtSearchValue = "";
                }

            }

            if (lblHis2 != "" && (strMode.ToLower() != "linkclick" && strMode.ToLower() != "search"))
            {
                strAutoart = lblHis2.Split('-')[1].ToString();
            }

            SqlConnection myConnection = new SqlConnection();
            myConnection = DBProc.getConnection(Session["sConnSiteDB"].ToString());
            myConnection.Open();
            SqlCommand sqlCmd = new SqlCommand("JobHistory", myConnection);

            if (strMode == "Excel")
            {
                sqlCmd = new SqlCommand("JobHistory_Excel", myConnection);

            }

            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.Parameters.Add(new SqlParameter("@strJbm_autoid", SqlDbType.VarChar)).Value = strCboVal;
            sqlCmd.Parameters.Add(new SqlParameter("@strSearchValue", SqlDbType.VarChar)).Value = strSearchValue;
            sqlCmd.Parameters.Add(new SqlParameter("@strAutoartid", SqlDbType.VarChar)).Value = strAutoart;
            sqlCmd.Parameters.Add(new SqlParameter("@TblArticleInfo", SqlDbType.VarChar)).Value = strArtcileinfo;
            sqlCmd.Parameters.Add(new SqlParameter("@TblStageInfo", SqlDbType.VarChar)).Value = strStageinfo;
            sqlCmd.Parameters.Add(new SqlParameter("@TblIssueInfo", SqlDbType.VarChar)).Value = strCustAcc + Init_Tables.gTblIssueInfo;
            sqlCmd.Parameters.Add(new SqlParameter("@TblAllocation", SqlDbType.VarChar)).Value = strCustAcc + Init_Tables.gTblJBM_Allocation;
            sqlCmd.Parameters.Add(new SqlParameter("@TblSplInst", SqlDbType.VarChar)).Value = strCustAcc + Init_Tables.gTblSplInstructions;
            sqlCmd.Parameters.Add(new SqlParameter("@TblProdStatus", SqlDbType.VarChar)).Value = strProdStatus;
            sqlCmd.Parameters.Add(new SqlParameter("@TblProdInfo", SqlDbType.VarChar)).Value = strProdinfo;
            sqlCmd.Parameters.Add(new SqlParameter("@issBased", SqlDbType.VarChar)).Value = strIss;
            sqlCmd.Parameters.Add(new SqlParameter("@strArtSearchValue", SqlDbType.VarChar)).Value = strArtSearchValue;

            SqlDataAdapter myReader = new SqlDataAdapter(sqlCmd);
            myReader.Fill(dsart);
            myConnection.Close();
            dsart.Tables[0].Columns.Add("Artwork", typeof(bool));
            for (int i = 0; i < dsart.Tables[0].Rows.Count; i++)
            {
                DataTable dt = new DataTable();

                if (strCustGroup == "CG001")
                {
                    dt = DBProc.GetResultasDataTbl("select ArtReq from " + Init_Tables.gTblRevProdInfo + " where autoartid='" + dsart.Tables[0].Rows[i]["JobID"].ToString() + "'", Session["sConnSiteDB"].ToString());
                }
                else
                {
                    dt = DBProc.GetResultasDataTbl("Select Art from " + Init_Tables.gTblJrnlInfo + " where JBM_AutoID = '" + dsart.Tables[0].Rows[i]["AutoId"].ToString() + "'", Session["sConnSiteDB"].ToString());
                }

                DataRow dr = dsart.Tables[0].Rows[i];

                if (dt.Rows.Count > 0)
                {
                    string Art = dt.Rows[0][0].ToString();
                    if ((Art == null || Art == "" || string.IsNullOrEmpty(Art)) || (Art == "False"))
                    {
                        dr["Artwork"] = false;
                    }
                    else
                    {
                        dr["Artwork"] = true;
                    }
                }
            }


            string query = "select Autoartid, revfinstage as Stage,Au_ReminderCancel,Ed_ReminderCancel,Pe_ReminderCancel,Pr_ReminderCancel from " + strCustAcc + Init_Tables.gTblStageInfo + " where autoartid='" + cboID + "' and revfinstage='fp'";
            DataTable dtreminder = new DataTable();
            dtreminder = DBProc.GetResultasDataTbl(query, Session["sConnSiteDB"].ToString());
            dtreminder.TableName = "Reminder";
            dsart.Tables.Add(dtreminder);

            string strQuery = "select p.AutoArtID, p.fprevptrinfo as Stage,p.OPS_wf as [Proof WF],p.mfrom as [Mail From],p.mto as [Author E-Mail],jbm_meditor as [Editor E-Mail], jbm_mpe as [PE E-Mail],PrEmail as [PR E-Mail],case p.CWP_Article when 0 then 'Normal' when 1 then 'CWP' when 2 then 'EPT' when 3 then 'SmartProof' when 4 then 'SmartPDF' end as [Proof Type] from " + strProdinfo + " p join " + strArtcileinfo + "  a on p.autoartid=a.autoartid join " + Init_Tables.gTblJrnlInfo + "  j on a.jbm_autoid= j.jbm_autoid where p.AutoArtID='" + cboID + "' order by Stage";
            DataTable dtsproof = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());
            dtsproof.TableName = "Proof";
            dsart.Tables.Add(dtsproof);

            query = "select Autoartid,FpRevPtrInfo as Stage,ToStop_Licencechaser from " + strProdinfo + " where AutoArtID='" + cboID + "'";
            DataTable dtlicence = new DataTable();
            dtlicence = DBProc.GetResultasDataTbl(query, Session["sConnSiteDB"].ToString());
            dtlicence.TableName = "Licence";
            dsart.Tables.Add(dtlicence);

            Session["dsart"] = dsart;
            Session["dsartNas"] = dsart;
            if (strMode == "Excel")
            {
                ExcelReport();
                return;
            }

            dtArticle = dsart.Tables[0];

            if (dtArticle.Rows.Count == 0)
            {
                //lblHis1.Text = "No Items Found";
                //divInternalnotes.Visible = false;
                //divArtNotes.Visible = false;
                //UpdatePanel1.Update();
                return;
            }
            else
            {
                Session["lblAutId"] = dtArticle.Rows[0]["AutoId"].ToString();
            }

            if (strCustAcc == "AC" || strCustAcc == "EM" || strCustAcc == "EH" || strCustAcc == "SB" || strCustAcc == "SM")
            {
                // btnGenerateTransmittal.Visible = true;

                if (strCustAcc == "AC")
                {
                    // btnUpdateCID.Visible = true;
                }
            }

            /* ***********************  */
            DataTable dtStage = dsart.Tables[1];
            dtArticleDetails = dsart.Tables[2];

            //divArtNotes.Visible = true;
            //divInternalnotes.Visible = true;
            //txtArtNotes.Text = Convert.ToString(dtArticle.Rows[0]["Notes"]);
            //txtInternalnotes.Text = Convert.ToString(dtArticle.Rows[0]["Remarks_Internal"]);

            Session["ArticleDetails"] = dtArticleDetails;

            //grdHis1.DataSource = dtArticle;
            //grdHis1.DataBind();

            if (strIss == "")  //if search article below columns should be hide  PDF_GEN_Start as [PDF Start],PDF_GEN_End as [PDF End]
            {
                //grdHis1.Columns[27].Visible = false;
                // grdHis1.Columns[28].Visible = false;
            }
            else  //if search issue below columns should be visible PDF_GEN_Start as [PDF Start],PDF_GEN_End as [PDF End]
            {
                // grdHis1.Columns[27].Visible = true;
                // grdHis1.Columns[28].Visible = true;
            }

            if (dtArticle.Rows.Count != 0)
            {
                //lblHis1.Text = "Article Details";
                // btnXlsExport.Visible = true;

                if (strCustAcc == "BK")
                {
                    // lblHis1.Text = "Chapter Details";

                }

            }

            ViewData["StageInfo"] = dsart.Tables[1];
            ViewData["AllocationInfo"] = dsart.Tables[2];
            if (dtArticle.Rows.Count == 1 || strAutoart != "" || strMode == "EditMode")
            {
                //grdHis2.DataSource = dtStage;
                //grdHis2.DataBind();

                if (dtStage.Rows.Count != 0 && dtArticle.Rows.Count > 1)
                {
                    //lblHis2.Text = "Stage Details -" + strAutoart;

                }
                else if (dtStage.Rows.Count != 0 && dtArticle.Rows.Count == 1)
                {
                    //lblHis2.Text = "Stage Details -" + dtArticle.Rows[0][0].ToString();

                }

                else if (strAutoart != "" && dtStage.Rows.Count == 0)
                {
                    //lblHis2.Text = " - Not available";

                }
            }
            string strIntrnlJID = "";
            strIntrnlJID = dsart.Tables[0].Rows[0][31].ToString();
            String strSelVolIssNo = ""; //For NAS Explorer
            strSelVolIssNo = dsart.Tables[0].Rows[0][8].ToString();    //For NAS Explorer
            string strSiteID = dsart.Tables[0].Rows[0]["SiteID"].ToString();
            string strCustSN = dsart.Tables[0].Rows[0]["CustSN"].ToString();
            string strJBMAID = dsart.Tables[0].Rows[0]["AutoId"].ToString();
            string sAutoArtID = dsart.Tables[0].Rows[0]["JobId"].ToString();
            string sInternalID = dsart.Tables[0].Rows[0]["InternalID"].ToString();
            string sArtPickupLoc = "";
            if (dsart.Tables[1].Rows.Count > 0)
                sArtPickupLoc = dsart.Tables[1].Rows[0]["ArtPickup_Loc"].ToString();

            if (strSiteID == "")
            {
                strSiteID = Session["sSiteID"].ToString();
            }

            if ((rdSearchList == "2" || (CustDisplayId == "1" && (strCustGroup != "CG002" && strCustAcc != "SA"))) && (cboID != ""))
            {
                DataTable dtIssueInfo = dsart.Tables[3];

                ViewData["IssueInfo"] = dsart.Tables[3];
                if (dtIssueInfo.Rows.Count != 0)
                {
                    // lblHis5.Text = "Issue details";
                    // divIssueDetails.Visible = true;
                    if (strCustAcc == "BK")
                    {
                        // lblHis5.Text = "Print Details";

                    }
                }

                else
                {
                    //if (grdHis1.Rows.Count != 0)
                    //{
                    //    lblHis5.Text = " - Not available";
                    //    divIssueDetails.Visible = false;
                    //    if (strCustAcc == "BK")
                    //    {
                    //        lblHis5.Text = "";

                    //    }
                    //}
                }

                //grdHis5.DataSource = dtIssueInfo;
                //grdHis5.DataBind();

            }
            if (lblHis2 != "")
            {
                //stageDiv.Visible = true;
            }


            //UpdatePanel1.Update();

        }
        public JsonResult AutoComplete(string JournalID, string txtsearch)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strQryText = "";
                string strJournal = "";
                if (JournalID != "" && JournalID != "0")
                {
                    strJournal = " and a.JBM_AutoID='" + JournalID + "' ";
                }

                if (txtsearch != "")
                {
                    strQryText = @"select AutoArtID as searchlst from " + strCustAcc + @"_ArticleInfo a where  AutoArtID like '" + txtsearch + @"%' " + strJournal  + " union select IntrnlID as searchlst from " + strCustAcc + @"_ArticleInfo a where  IntrnlID like '" + txtsearch + @"%' " + strJournal + " union select ChapterID as searchlst from " + strCustAcc + @"_ArticleInfo a where  ChapterID like '" + txtsearch + @"%'" + strJournal + " union select cadaid as searchlst from " + strCustAcc + @"_ArticleInfo a where  cadaid like '" + txtsearch + @"%' " + strJournal + "";

                    //                    strQryText = @"select AutoArtID as searchlst from " + strCustAcc + @"_ArticleInfo a join JBM_Info ji on 
                    //a.JBM_AutoID=ji.JBM_AutoID where  AutoArtID like '" + txtsearch + @"%' union 
                    //select IntrnlID as searchlst from " + strCustAcc + @"_ArticleInfo a join JBM_Info ji on 
                    //a.JBM_AutoID=ji.JBM_AutoID where  IntrnlID like '" + txtsearch + @"%' union 
                    //select ChapterID as searchlst from " + strCustAcc + @"_ArticleInfo a join JBM_Info ji on 
                    //a.JBM_AutoID=ji.JBM_AutoID where  ChapterID like '" + txtsearch + @"%' union 
                    //select cadaid as searchlst from " + strCustAcc + @"_ArticleInfo a join JBM_Info ji on 
                    //a.JBM_AutoID=ji.JBM_AutoID where  cadaid like '" + txtsearch + @"%'";

                    DataTable dt = new DataTable();
                    dt = DBProc.GetResultasDataTbl(strQryText, Session["sConnSiteDB"].ToString());

                    var jsonString = from a in dt.AsEnumerable() select new[] { a[0].ToString() };
                    return Json(jsonString, JsonRequestBehavior.AllowGet);
                }
                return Json("Success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("Error", JsonRequestBehavior.AllowGet);
            }
        }
        public string CreateDeleteBtn(string uniqueID, string Iss, string stage)
        {
            string formControl = string.Empty;
            try
            {
                if (uniqueID != "")
                {
                    formControl = "<a id = 'btndelete" + uniqueID + "' href='javascript:void(0);' onclick=\"DeleteIssueAllocationDetails('" + uniqueID + "','" + Iss + "','" + stage + "');\"><i class='far fa-trash-alt' style='color:#eb2227;font-size:14px'></i></a>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult GetAllocation(string cboID, string cboIssues)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strQuery = "Select a.iss,a.Stage,a.allocatedto as EmpId,a.allocatedtoname as [Allocated To], convert(varchar(12),cast(a.AllocatedDate as DateTime),106) as [Allocated Date], a.Pages, b.DeptSN as Dept, a.allocatedbyname as [Allocated By] from " + strCustAcc + "_Allocation a,JBM_DepartmentMaster b where a.DeptCode=b.DeptCode and a.AutoArtID='" + cboID + "' and a.iss='" + cboIssues + "' order by Stage";

                DataTable dtAllot = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());


                var JSONString = from a in dtAllot.AsEnumerable()
                                 select new[] {
                                     CreateDeleteBtn(a[2].ToString(),a[0].ToString(),a[1].ToString()),
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetProduction(string cboID, string cboIssues)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdStatus = strCustAcc + Init_Tables.gTblProdStatus;
                string strQuery = "SELECT [Emp. Name],Activity,GP,StageDesc,K.Stage,[Corr. Cnt],[Emp In-time],[Emp Out-Time],Remarks FROM (SELECT Allocatedtoname AS [Emp. Name], PD.StageDesc AS Activity, p.gprocess AS GP,CASE CHARINDEX('_', ArtStageInfo) WHEN '0' THEN ArtStageInfo ELSE SUBSTRING(p.ArtStageInfo, 1, CHARINDEX('_', p.ArtStageInfo)-1) END AS Stage, CASE CHARINDEX('_', ArtStageInfo) WHEN '0' THEN 'Corr_0' ELSE 'Corr_'+SUBSTRING(p.ArtStageInfo, CHARINDEX('_', p.ArtStageInfo)+1, 10) END AS [Corr. Cnt], CONVERT(varchar, EmpIntime, 106)+' '+SUBSTRING(CONVERT(varchar(20), EmpIntime, 22), 10, 11) AS [Emp In-time], EmpIntime, CONVERT(varchar, EmpOutTime, 106)+' '+SUBSTRING(CONVERT(varchar(20), EmpOutTime, 22), 10, 11) AS [Emp Out-Time], p.Remarks, p.ArtStageCurrStat FROM JBM_ProdArtStatDesDept AS PD INNER JOIN " + strProdStatus + " AS P ON p.EmpArtStage = PD.DeptActivity WHERE p.iss='" + cboIssues + "' and AutoArtID = '" + cboID + "') AS K INNER JOIN JBM_ProdArtStatDesDept AS P_1 ON K.ArtStageCurrStat = p_1.DeptActivity ORDER BY [EmpIntime] ASC;";

                DataTable dtProd = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());


                var JSONString = from a in dtProd.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString(),
                                     a[8].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetHold(string cboID, string cboIssues)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdStatus = strCustAcc + Init_Tables.gTblProdStatus;
                string strQuery = "select (case status when '0' then 'Released' else 'Hold' end) as HoldStatus,(select empsurname from jbm_employeemaster where q.mailsentby=empautoid) as MS,(CONVERT(VARCHAR,HoldDate,106) + ' ' + substring(CONVERT(VARCHAR(20), HoldDate, 22),10,11)) as [Hold_Date], CustAccess,AutoArtID,RevFinStage,HoldBy,ReleasedBy,QryTo,HoldDate,ReasonforHold,MailFrom,MailTo,MailCc,MailBcc,MailSubject,MailContent,MailSentBy,MailQryDate,LastFollowUpDate,FollowUpDates,ReleaseDate,ReleaseComments,Status,QryMail_Status,JBM_Location,MailCount,iss from  " + Init_Tables.gTblJBM_QryInfo + " q where AutoArtID='" + cboID + "' order by holddate desc";

                DataTable dtProd = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtProd.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult GetHold2(string cboID)
        {
            try
            {
                string strQuery = "select (case status when '0' then 'Released' else 'Hold' end) as HoldStatus,(select empsurname from jbm_employeemaster where q.mailsentby=empautoid) as MS,(CONVERT(VARCHAR,HoldDate,106) + ' ' + substring(CONVERT(VARCHAR(20), HoldDate, 22),10,11)) as [Hold_Date], CustAccess,AutoArtID,RevFinStage,HoldBy,ReleasedBy,QryTo,HoldDate,ReasonforHold,MailFrom,MailTo,MailCc,MailBcc,MailSubject,MailContent,MailSentBy,MailQryDate,LastFollowUpDate,FollowUpDates,ReleaseDate,ReleaseComments,Status,QryMail_Status,JBM_Location,MailCount,iss from  " + Init_Tables.gTblJBM_QryInfo + " q where AutoArtID='" + cboID + "' order by holddate desc";

                DataTable dtHold = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtHold.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetSplInstructions(string cboID, string cboIssues)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdStatus = strCustAcc + Init_Tables.gTblProdStatus;
                string strSplInstQuery = "select spl.Stage as Stage,spl.Instruction as Instruction,convert(varchar(12),spl.InstDate,106) + ' ' +  convert(varchar(8),spl.InstDate,108) as InstDate from " + strCustAcc + Init_Tables.gTblSplInstructions + " spl where spl.autoartid='" + cboID + "'  and spl.iss='" + cboIssues + "'";
                if (strCustAcc == "TF")
                {
                    strSplInstQuery = "select s.Stage as Stage, (case ISNUMERIC(SUBSTRING(s.Instruction,1,3)) when 0 then s.Instruction else case ISNUMERIC(SUBSTRING(s.Instruction,1,1)) when 1 then (select Instructions from " + strCustAcc + Init_Tables.gTblSplInstructions + " where s.Instruction= autoseqid)else s.Instruction  end end) as Instruction,convert(varchar(12),s.InstDate,106) + ' ' +  convert(varchar(8),s.InstDate,108) as InstDate from " + strCustAcc + Init_Tables.gTblSplInstructions + " s where autoartid='" + cboID + "' and iss='" + cboIssues + "'";
                }
                DataTable dtsubSplInst = DBProc.GetResultasDataTbl(strSplInstQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtsubSplInst.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetSplInstructions2(string cboID)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strSplInstQuery = "select spl.Stage as Stage,spl.Instruction as Instruction,convert(varchar(12),spl.InstDate,106) + ' ' +  convert(varchar(8),spl.InstDate,108) as InstDate from " + strCustAcc + Init_Tables.gTblSplInstructions + " spl where spl.autoartid='" + cboID + "'";

                if (strCustAcc == "TF")
                {
                    strSplInstQuery = "select s.Stage as Stage, (case ISNUMERIC(SUBSTRING(s.Instruction,1,3)) when 0 then s.Instruction else case ISNUMERIC(SUBSTRING(s.Instruction,1,1)) when 1 then (select Instructions from " + strCustAcc + Init_Tables.gTblSplInstructions + " where s.Instruction= autoseqid)else s.Instruction  end end) as Instruction,convert(varchar(12),s.InstDate,106) + ' ' +  convert(varchar(8),s.InstDate,108) as InstDate from " + strCustAcc + Init_Tables.gTblSplInstructions + " s where autoartid='" + cboID + "'";
                }

                DataTable dtsubSplInst = DBProc.GetResultasDataTbl(strSplInstQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtsubSplInst.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        public string CreateChkInp(string uniqueID, string Input, string stage)
        {
            string formControl = string.Empty;
            try
            {
                if (uniqueID != "")
                {
                    if (Input == "true")
                        formControl = "<input class='pt-3' type='checkbox' id='" + stage.Trim() + "_" + uniqueID + "' style='width:12px' disabled checked>";
                    else
                        formControl = "<input class='pt-3' type='checkbox' id='" + stage.Trim() + "_" + uniqueID + "' style='width:12px' disabled unchecked>";

                    //formControl = "<a id = 'btndelete" + uniqueID + "' href='javascript:void(0);' onclick=\"DeleteIssueAllocationDetails('" + uniqueID + "','" + Iss + "','" + stage + "');\"><i class='far fa-trash-alt' style='color:#eb2227;font-size:14px'></i></a>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult GetStageOther(string cboID)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strStageinfo = strCustAcc + Init_Tables.gTblStageInfo;
                //string strQuery = "select revfinstage as Stage, convert(char(15),ceRecdate,106) as [CE Recd. Date], convert(char(15),ceDuedate,106) as [CE Due Date], convert(char(15),ceDispDate,106) as [CE Disp. Date], convert(char(15), CrsTypeAutRec,106) as [Aut Corr], convert(char(15), CrsTypeEdRec,106) as [Ed Corr], crstypePeRec as [PE Corr], crstypeProofReaderRec as [PR Corr] , Edt_corr_Appr as [Edt Approve], PE_corr_Appr as [PE Approve], PR_corr_Appr as [PR_Corr_Appr], Aut_Corr_Appr, CorrFigs as [Corr. Figs] from " + strStageinfo + " where AutoArtID='" + cboID + "' order by revfinstage";
                string strQuery = @"select revfinstage as Stage, convert(char(15),ceRecdate,106) as [CE Recd. Date], convert(char(15),ceDuedate,106) as [CE Due Date], 
                convert(char(15),ceDispDate,106) as [CE Disp. Date], convert(char(15), s.CrsTypeAutRec,106) as [Aut Corr], convert(char(15), s.CrsTypeEdRec,106) 
                as [Ed Corr], crstypePeRec as [PE Corr], crstypeProofReaderRec as [PR Corr] , Edt_corr_Appr as [Edt Approve], PE_corr_Appr as [PE Approve], 
                PR_corr_Appr as [PR_Corr_Appr], Aut_Corr_Appr, CorrFigs as [Corr. Figs],case when Au_ReminderCancel=1 then 'true' else 'false' end as Au_ReminderCancel,
case when Ed_ReminderCancel=1 then 'true' else 'false' end as Ed_ReminderCancel,
case when Pe_ReminderCancel=1 then 'true' else 'false' end as Pe_ReminderCancel ,
case when Pr_ReminderCancel=1 then 'true' else 'false' end as Pr_ReminderCancel,
case when ToStop_Licencechaser=1 then 'true' else 'false' end as ToStop_Licencechaser,s.AutoArtID from " + strStageinfo + @" s join " + strProdinfo + @" p on s.AutoArtID=p.AutoArtID where s.AutoArtID='" + cboID + @"' and revfinstage='fp' 
                order by revfinstage";
                DataTable dtst = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtst.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString(),
                                     a[8].ToString(),
                                     a[9].ToString(),
                                     a[10].ToString(),
                                     a[11].ToString(),
                                     a[12].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetProof(string cboID)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strStageinfo = strCustAcc + Init_Tables.gTblStageInfo;
                string strQuery = @"select revfinstage as Stage,p.OPS_wf as [Proof WF],p.mfrom as [Mail From],p.mto as [Author E-Mail],jbm_meditor as [Editor E-Mail], jbm_mpe as [PE E-Mail],PrEmail as [PR E-Mail],case p.CWP_Article when 0 then 'Normal' when 1 then 'CWP' when 2 then 'EPT' when 3 then 'SmartProof' when 4 then 'SmartPDF' end as [Proof Type], case when Au_ReminderCancel=1 then 'true' else 'false' end as Au_ReminderCancel,
case when Ed_ReminderCancel=1 then 'true' else 'false' end as Ed_ReminderCancel,
case when Pe_ReminderCancel=1 then 'true' else 'false' end as Pe_ReminderCancel ,
case when Pr_ReminderCancel=1 then 'true' else 'false' end as Pr_ReminderCancel,
case when ToStop_Licencechaser=1 then 'true' else 'false' end as ToStop_Licencechaser,s.AutoArtID from " + strStageinfo + @" s join " + strProdinfo + @" p on s.AutoArtID=p.AutoArtID 
join " + strArtcileinfo + "  a on p.autoartid=a.autoartid join JBM_Info  j on a.jbm_autoid= j.jbm_autoid where s.AutoArtID='" + cboID + @"' 
                order by revfinstage";
                DataTable dtsproof = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtsproof.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString(),
                                     CreateChkInp(a[13].ToString(),a[8].ToString(),"chkAuRC"),
                                     CreateChkInp(a[13].ToString(),a[9].ToString(),"chkEdRC"),
                                     CreateChkInp(a[13].ToString(),a[10].ToString(),"chkPeRC"),
                                     CreateChkInp(a[13].ToString(),a[11].ToString(),"chkPrRC"),
                                     CreateChkInp(a[13].ToString(),a[12].ToString(),"chkstop")
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public string CreateBtnforBindReminder(string uniqueID)
        {
            string formControl = string.Empty;
            try
            {
                if (uniqueID != "")
                {
                    formControl = "<a class='Edit' href='javascript:;' onclick=\"EditBindReminder('" + uniqueID + "');\"><i class='far fa-edit' style='color:#28a745;font-size:14px'></i></a>";
                    formControl += "<a class='Update' href='javascript:;' onclick=\"UpdateBindReminder('" + uniqueID + "');\" style='display:none'><i class='far fa-save' style='color:#28a745;font-size:14px'></i></a>";
                    formControl += "<a class='Cancel' href='javascript:;' onclick=\"CancelBindReminder('" + uniqueID + "');\" style='display:none'><i class='far fa-window-close' style='color:#eb2227;font-size:14px'></i></a>";
                }
                return formControl;
            }
            catch (Exception)
            {
                return "";
            }
        }
        [SessionExpire]
        public ActionResult BindReminder(string cboID)
        {

            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string query = "select Autoartid, revfinstage as Stage,Au_ReminderCancel,Ed_ReminderCancel,Pe_ReminderCancel,Pr_ReminderCancel from " + strCustAcc + Init_Tables.gTblStageInfo + " where autoartid='" + cboID + "' and revfinstage='fp'";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(query, Session["sConnSiteDB"].ToString());
                DataSet ds = (DataSet)Session["dsart"];
                var JSONString = from a in dt.AsEnumerable()
                                 select new[] {CreateBtnforBindReminder(a[0].ToString()),
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }


        [SessionExpire]
        public ActionResult BindLicenceGrid(string cboID)
        {

            try
            {
                string query = "select Autoartid,FpRevPtrInfo as Stage,ToStop_Licencechaser from " + Init_Tables.gTblRevProdInfo + " where AutoArtID='" + cboID + "'";

                DataTable dt = new DataTable();
                dt = DBProc.GetResultasDataTbl(query, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dt.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetAllocation2(string cboID)
        {
            try
            {
                string strQuery = "Select a.iss,a.Stage,a.allocatedto as EmpId,a.allocatedtoname as [Allocated To], convert(varchar(12),cast(a.AllocatedDate as DateTime),106) as [Allocated Date], a.Pages, b.DeptSN as Dept, a.allocatedbyname as [Allocated By] from " + Session["sCustAcc"].ToString() + Init_Tables.gTblJBM_Allocation + " a," + Init_Tables.gTblDepartment + " b where a.DeptCode=b.DeptCode and AutoArtID='" + cboID + "' order by Stage";

                DataTable dtAllot = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtAllot.AsEnumerable()
                                 select new[] {
                                     CreateDeleteBtn(a[2].ToString(),a[0].ToString(),a[1].ToString()),
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult GetProduction2(string cboID)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdStatus = strCustAcc + Init_Tables.gTblProdStatus;
                string strQuery = "SELECT [Emp. Name],Activity,GP,StageDesc,K.Stage,[Corr. Cnt],[Emp In-time],[Emp Out-Time],Remarks FROM (SELECT Allocatedtoname AS [Emp. Name], PD.StageDesc AS Activity, p.gprocess AS GP,CASE CHARINDEX('_', ArtStageInfo) WHEN '0' THEN ArtStageInfo ELSE SUBSTRING(p.ArtStageInfo, 1, CHARINDEX('_', p.ArtStageInfo)-1) END AS Stage, CASE CHARINDEX('_', ArtStageInfo) WHEN '0' THEN 'Corr_0' ELSE 'Corr_'+SUBSTRING(p.ArtStageInfo, CHARINDEX('_', p.ArtStageInfo)+1, 10) END AS [Corr. Cnt], CONVERT(varchar, EmpIntime, 106)+' '+SUBSTRING(CONVERT(varchar(20), EmpIntime, 22), 10, 11) AS [Emp In-time], EmpIntime, CONVERT(varchar, EmpOutTime, 106)+' '+SUBSTRING(CONVERT(varchar(20), EmpOutTime, 22), 10, 11) AS [Emp Out-Time], p.Remarks, p.ArtStageCurrStat FROM JBM_ProdArtStatDesDept AS PD INNER JOIN " + strProdStatus + " AS P ON p.EmpArtStage = PD.DeptActivity WHERE AutoArtID = '" + cboID + "') AS K INNER JOIN JBM_ProdArtStatDesDept AS P_1 ON K.ArtStageCurrStat = p_1.DeptActivity ORDER BY [EmpIntime] ASC;";
                DataTable dtProd = DBProc.GetResultasDataTbl(strQuery, Session["sConnSiteDB"].ToString());

                var JSONString = from a in dtProd.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString(),
                                     a[7].ToString(),
                                     a[8].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult GetTrack(string cboID)
        {
            try
            {
                string strCustAccess = Session["sCustAcc"].ToString();
                string strSplInstQuery = "Select E.EmpLogin as [Login], E.EmpName as [Name], D.DeptSN as [Dept.], (CONVERT(VARCHAR,AccTime,106) + ' ' + substring(CONVERT(VARCHAR(20), AccTime, 22),10,11)) as [Time], P.Descript as [Task],P.ShortDescript as [ShortDesc], P.LinkTo as [Link] from " + Init_Tables.gTblProdAccess + " P, " + Init_Tables.gTblEmployee + " E, " + Init_Tables.gTblDepartment + " D where E.EmpAutoID = P.EmpAutoID and D.DeptCode = E.DeptCode and P.AutoArtID = '" + cboID + "' and P.CustAcc = '" + strCustAccess + "' order by [AccTime] asc";

                DataTable dtsubSplInst = DBProc.GetResultasDataTbl(strSplInstQuery, Session["sConnSiteDB"].ToString());
                var JSONString = from a in dtsubSplInst.AsEnumerable()
                                 select new[] {
                                     a[0].ToString(),
                                     a[1].ToString(),
                                     a[2].ToString(),
                                     a[3].ToString(),
                                     a[4].ToString(),
                                     a[5].ToString(),
                                     a[6].ToString()
                 };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        public void drpSelected_change(string QryStrAutoid)
        {
            // FolderDir.Visible = true;
            //lblJid.Text = "";
            //Clear();
            //txtID.Text = "";
            if (QryStrAutoid != null && QryStrAutoid != "")
            {
                //This string replaced because to maintain the current value while redirects back from the holdarticle page.
                string CurrentValue = "";
                if (Request.QueryString["Autoid"] != null)
                {
                    CurrentValue = Request.QueryString["Autoid"];
                    QryStrAutoid = QryStrAutoid.Replace(CurrentValue, "");
                    //cboID.ClearSelection();
                }
                //txtID.Text = QryStrAutoid;

            }


            //UpdatePanel1.Update();

        }
        public void chkbilled_change(string chkBilling)
        {
            //Clear();
            //txtID.Text = string.Empty;
            //cboID.ClearSelection();
            //lblJid.Text = string.Empty;
            //UpdatePanel1.Update();

            if (chkBilling == "1")
            {
                //Change for billed DB

                strStageinfo = "[" + Session["sConnSiteDB"].ToString() + "].[dbo]." + Init_Tables.gTblStageInfo;
                strArtcileinfo = "[" + Session["sConnSiteDB"].ToString() + "].[dbo]." + Init_Tables.gTblChapterOrArticleInfo;
                strProdinfo = "[" + Session["sConnSiteDB"].ToString() + "].[dbo]." + Init_Tables.gTblRevProdInfo;
                strProdStatus = "[" + Session["sConnSiteDB"].ToString() + "].[dbo]." + Init_Tables.gTblProdStatus;
                //Change for billed DB
            }
        }
        [SessionExpire]
        public ActionResult cboCust_SelectedIndexChanged(string cboCust)
        {
            try
            {
                if (Session["LoadJournal"] == null)

                {
                    string strCustAcc = Session["sCustAcc"].ToString();
                    string strjournalQuery = string.Empty;
                    // Clear();
                    //txtID.Text = "";
                    if (cboCust != "" || cboCust != null)
                    {
                        Session["CustSelect"] = cboCust;
                        string strQueryFinal = "Select JBM_Intrnl,JBM_AutoID from " + Init_Tables.gTblJrnlInfo + " where JBM_AutoID like '%" + strCustAcc + "%' and CustId='" + cboCust + "' and JBM_Disabled='0' order by JBM_ID asc";

                        DataSet ds = new DataSet();
                        ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                        var JSONString = from a in ds.Tables[0].AsEnumerable()
                                         select new[] {a[0].ToString(),a[1].ToString()
                                };
                        return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    Session["LoadJournal"] = null;
                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult LoadJournal()
        {
            try
            {
                Session["LoadJournal"] = "LoadJournal";
                string strCustAcc = Session["sCustAcc"].ToString();
                string strjournalQuery = string.Empty;
                string strQueryFinal = "Select JBM_Intrnl,JBM_AutoID from " + Init_Tables.gTblJrnlInfo + " where JBM_AutoID like '%" + strCustAcc + "%' and JBM_Disabled='0' order by JBM_ID asc";

                DataSet ds = new DataSet();
                ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                var JSONString = from a in ds.Tables[0].AsEnumerable()
                                 select new[] {a[0].ToString(),a[1].ToString()
                                };
                return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        [SessionExpire]
        public ActionResult cboIssues_SelectedIndexChanged(string strAutoart, string strMode, string txtID, string cboID, string rdSearchList, string cboIssues)
        {

            if (cboID != "")
            {
                GetHistoryLink(strAutoart, strMode, txtID, cboID, rdSearchList, cboIssues, "", "");
                Session["List"] = "ShowIss";
                Session["JIDSelect"] = cboID;
                Session["IssueSelect"] = cboIssues;
                Session["optionISS"] = "cboIssues";
                //lblHis2.Text = "";
                //lblHis3.Text = "";
                //lblHis4.Text = "";
                //lblHd.Text = "";
            }
            //stageDiv.Visible = false;
            //grdHis2.DataSource = "";
            //grdHis2.DataBind();
            //grdHis3.DataSource = "";
            //grdHis3.DataBind();
            //grdHis4.DataSource = "";
            //grdHis4.DataBind();
            //grdHold.DataSource = "";
            //grdHold.DataBind();

            //BtnAllotiss.Visible = true;
            //BtnSpliss.Visible = true;
            //BtnHoldiss.Visible = true;
            //BtnProdiss.Visible = true;

            //lblSpliss.Text = "";
            //lblHdiss.Text = "";
            //lblHis3iss.Text = "";
            //lblHis4iss.Text = "";
            //lblAllocMessiss.Text = "";

            //grdHis3iss.DataSource = "";
            //grdHis3iss.DataBind();
            //grdHis4iss.DataSource = "";
            //grdHis4iss.DataBind();
            //grdHoldiss.DataSource = "";
            //grdHoldiss.DataBind();
            //RptrSplInstiss.DataSource = "";
            //RptrSplInstiss.DataBind();            
            //return View(dsart);
            string result = "";
            if (Session["List"] != null)
                result = "Success";
            else
                result = "Fail";


            return Json(new { dataComp = result }, JsonRequestBehavior.AllowGet);

        }

        [SessionExpire]
        public ActionResult GetISSList(string sJID)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                if (sJID != "" || sJID != null)
                {
                    string strQueryFinal = "select distinct a.iss as Issues from " + strCustAcc + Init_Tables.gTblChapterOrArticleInfo + " a join " + Init_Tables.gTblJrnlInfo + " j on a.jbm_autoid=j.jbm_autoid where j.jbm_autoid='" + sJID + "' And a.iss Is Not null";

                    DataSet ds = new DataSet();
                    ds = DBProc.GetResultasDataSet(strQueryFinal, Session["sConnSiteDB"].ToString());

                    var JSONString = from a in ds.Tables[0].AsEnumerable()
                                     select new[] {a[0].ToString()
                                };
                    return Json(new { dataComp = JSONString }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { dataComp = "NoData" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { data = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult cboID_SelectedIndexChanged(string CustDisplayId, string cboID, string chkBilling, string cboIssues, string rdSearchList)
        {
            string strCustGroup = Session["sCustGroup"].ToString();
            string strCustAccess = Session["sCustAcc"].ToString();
            // trCIDupdate.Visible = false;
            if (cboID != null && cboID != "")
            {
                //lblJid.Text = "(" + cboID + ")";

                if (chkBilling == "1")
                {
                    string strQueryIssues = "select distinct a.iss as Issues from " + Session["sConnSiteDB"].ToString() + ".dbo." + Init_Tables.gTblChapterOrArticleInfo + " a join " + Init_Tables.gTblJrnlInfo + " j on a.jbm_autoid=j.jbm_autoid where j.jbm_autoid='" + cboID + "' And a.iss Is Not null";
                    List<SelectListItem> lstJID = new List<SelectListItem>();
                    DataSet dsJID = new DataSet();
                    dsJID = DBProc.GetResultasDataSet(strQueryIssues, Session["sConnSiteDB"].ToString());
                    if (dsJID.Tables[0].Rows.Count > 0)
                    {
                        for (int intCount = 0; intCount < dsJID.Tables[0].Rows.Count; intCount++)
                        {
                            string strEmpAutoID = dsJID.Tables[0].Rows[intCount]["iss"].ToString();
                            string strEmpName = dsJID.Tables[0].Rows[intCount]["iss"].ToString();
                            lstJID.Add(new SelectListItem
                            {
                                Text = strEmpName.ToString(),
                                Value = strEmpAutoID.ToString()
                            });
                        }

                    }

                    ViewBag.Isslist = lstJID;
                }

                if (rdSearchList == "2")
                {
                    // AjaxScriptManager.SetFocus(cboIssues);
                }
                else if (CustDisplayId == "1")
                {
                    if (strCustGroup == "CG002" || strCustAccess == "SA")
                    {
                        //Clear();
                        //txtID.Text = "";
                        //UpdatePanel1.Update();
                    }
                    else
                    {
                        GetHistoryLink("", "", "AJ013", "AJ013", "1", "--No Issues Found", "", "");
                    }
                }

            }

            else
            {

                //cboIssues.Items.Clear();
                //cboIssues.Items.Add(new ListItem("", ""));//(New ListItem("", ""));

            }

            if (CustDisplayId != "1")
            {
                //Clear();
                //txtID.Text = "";
                //UpdatePanel1.Update();

            }

            {
                //if (grdHis1.Rows.Count == 1)
                //{
                //    stageDiv.Visible = true;
                //}
                //else { stageDiv.Visible = false; }


                //UpdatePanel1.Update();
            }
            return View(dsart);
        }
        [SessionExpire]
        public ActionResult txtID_TextChanged(string strAutoart, string strMode, string txtID, string cboID, string rdSearchList, string cboIssues, string lblHis2, string isBookJnl)
        {
            //Clear();
            if (txtID != "")
            {
                GetHistoryLink("", strMode, txtID, "", rdSearchList, cboIssues, "", "");
                Session["List"] = "ShowTxt";
                Session["JIDSelect"] = cboID;
                Session["IssueSelect"] = cboIssues;
                Session["optionISS"] = "txtID";
            }
            //trCIDupdate.Visible = false;
            //BtnHold.Visible = true;
            //BtnSpl.Visible = true;
            //btnTrack.Visible = true;

            //UpdatePanel1.Update();
            string result = "";
            if (Session["List"] != null)
                result = "Success";
            else
                result = "Fail";


            return Json(new { dataComp = result }, JsonRequestBehavior.AllowGet);
        }
        [SessionExpire]
        public ActionResult lnkJob_Click(string strJbid)
        {
            try
            {
                GetHistoryLink(strJbid, "linKClick", "", "", "", "", "", "");
                Session["List"] = "ShowIss";
                //Session["JIDSelect"] = cboID;
                //Session["IssueSelect"] = cboIssues;
                string result = "";
                if (Session["List"] != null)
                    result = "Success";
                else
                    result = "Fail";


                return Json(new { dataComp = result }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [SessionExpire]
        public ActionResult UpdateArticleDetails(string VendorPgs, string Art, string lblAutId, string sjobid, string txtNoFigs, string txtNoNRefs, string textMSP, string textTSP, string textTables, string txtVendorPgs, string txtMScriptId, string txtDOI)
        {
            try
            {
                UpdateArticleDetails1(VendorPgs, Art, lblAutId, sjobid, txtNoFigs, txtNoNRefs, textMSP, textTSP, textTables, txtVendorPgs, txtMScriptId, txtDOI);
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        public void UpdateArticleDetails1(string VendorPgs, string Art, string lblAutId, string jobid, string txtNoFigs, string txtNoNRefs, string textMSP, string textTSP, string textTables, string txtVendorPgs, string txtMScriptId, string txtDOI)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strQuery = "";
                string strCustGroup = Session["sCustGroup"].ToString();
                if (strCustGroup == "CG001")
                {

                    if ((Art == null || Art == "" || string.IsNullOrEmpty(Art)) || (Art == "0"))
                    {
                        Art = "0";
                    }
                    else
                    {
                        Art = "1";
                    }


                    if (Art == "1")
                    {
                        if (txtNoFigs == "0" || txtNoFigs == string.Empty || txtNoFigs.ToLower() == "null")
                        {
                            //lblArtMess.Text = "Required specfic figure number";
                            return;
                        }

                    }

                    if (Art == "0")
                    {
                        int Nofigs;
                        bool isInt = int.TryParse(txtNoFigs, out Nofigs);

                        if (isInt)
                        {
                            if (Nofigs > 0)
                            {
                                // lblArtMess.Text = "Art Not Received,allow Zero Only";
                                return;
                            }
                        }
                        else
                        {
                            if (txtNoFigs.Trim() != "")
                            {
                                // lblArtMess.Text = "NumOfFigure not allowed non-numeric";
                                return;
                            }
                        }
                    }
                    // RecordManager.UpdateRecord("Update " + Init_Tables.gTblRevProdInfo + " set ArtReq=" + Art + " where autoartid='" + jobid + "'", "");
                    strQuery = DBProc.GetResultasString("Update " + strCustAcc + Init_Tables.gTblRevProdInfo + " set ArtReq=" + Art + " where autoartid='" + jobid + "'", Session["sConnSiteDB"].ToString());
                }

                if (textTSP != "")
                {
                    textTSP = "'" + textTSP + "'";
                }
                else
                {

                    textTSP = "Null";
                }

                if (textMSP != "")
                {
                    textMSP = "'" + textMSP + "'";
                }
                else
                {
                    textMSP = "Null";
                }
                if (txtNoFigs != "")
                {
                    txtNoFigs = "'" + txtNoFigs + "'";
                }
                else
                {
                    txtNoFigs = "Null";
                }

                if (txtNoNRefs != "")
                {
                    txtNoNRefs = "'" + txtNoNRefs + "'";
                }
                else
                {
                    txtNoNRefs = "Null";
                }

                string venPgs = string.Empty;

                if (txtVendorPgs != "")
                {
                    venPgs = txtVendorPgs;
                    txtVendorPgs = "'" + txtVendorPgs + "'";

                }
                else
                {

                    txtVendorPgs = "Null";
                }

                if (textTables != "")
                {
                    textTables = "'" + textTables + "'";
                }
                else
                {
                    textTables = "Null";
                }

                if (txtDOI != "")
                {
                    txtDOI = "'" + txtDOI + "'";
                }
                else
                {
                    txtDOI = "Null";
                }

                if (txtMScriptId != "")
                {
                    txtMScriptId = "'" + txtMScriptId + "'";
                }
                else
                {
                    txtMScriptId = "Null";
                }
                strQuery = "update " + strArtcileinfo + " set NumofMSP=" + textMSP + ", NumofFigures=" + txtNoFigs + ", NumofRefs=" + txtNoNRefs + ", ActualPages=" + textTSP + ",NumofTables=" + textTables + ",VendorPgs=" + txtVendorPgs + ",ManuscriptID=" + txtMScriptId + ",DOI=" + txtDOI + " where autoartid='" + jobid + "' ";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                if (VendorPgs != venPgs)
                {

                    string strBillingExists = "select TatAmnt,ErrorDeduction,RatesPerPage,vtype from " + strCustAcc + "_prodstatus where autoartid='" + jobid + "' and vtype is not null and empouttime is not null";
                    DataTable billdata = DBProc.GetResultasDataTbl(strBillingExists, Session["sConnSiteDB"].ToString());
                    if (billdata != null)
                    {

                        decimal intTatamount = 0;
                        decimal BasicRate = 0;
                        decimal ErrorDeduc = 0;

                        if (!DBNull.Value.Equals(billdata.Rows[0]["TatAmnt"]))
                        {

                            intTatamount = Convert.ToDecimal(billdata.Rows[0]["TatAmnt"]);

                        }
                        if (!DBNull.Value.Equals(billdata.Rows[0]["RatesPerPage"]))
                        {

                            BasicRate = Convert.ToDecimal(billdata.Rows[0]["RatesPerPage"]);

                        }
                        if (!DBNull.Value.Equals(billdata.Rows[0]["ErrorDeduction"]))
                        {

                            ErrorDeduc = Convert.ToDecimal(billdata.Rows[0]["ErrorDeduction"]);

                        }

                        decimal billiingRate = Convert.ToDecimal(Convert.ToDecimal(txtVendorPgs) * ((BasicRate + intTatamount) - ErrorDeduc));


                        strBillingExists = "Update " + strCustAcc + "_prodstatus set BillingRate=" + billiingRate + ",PgsCompleted='" + txtVendorPgs + "' where AutoArtID='" + jobid + "' and ArtStageInfo='FP' and vtype is not null and empouttime is not null";

                        if (billdata.Rows[0]["vtype"].ToString().ToLower() == "fl" || billdata.Rows[0]["vtype"].ToString().ToLower() == "wfh")
                        {
                            strQuery = DBProc.GetResultasString(strBillingExists, Session["sConnSiteDB"].ToString());
                        }

                    }

                }

                strQuery = DBProc.GetResultasString("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + strCustAcc + "', '" + jobid + "' , '" + Session["UserID"] + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','ArticleDetailsUpdated in Jobhistory','JobHistory','Null')", Session["sConnSiteDB"].ToString());

            }
            catch (Exception ex)
            {
            }
        }
        public void UpdateEPTReminderDetails(string sAuRC, string sEdRC, string sjobid, string sPeRC, string sPrRC)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strQuery = "";
                if (sAuRC != "")
                {
                    sAuRC = "" + sAuRC + "";
                }
                else
                {
                    sAuRC = "Null";
                }
                if (sEdRC != "")
                {
                    sEdRC = "" + sEdRC + "";
                }
                else
                {
                    sEdRC = "Null";
                }
                if (sPeRC != "")
                {
                    sPeRC = "" + sPeRC + "";
                }
                else
                {
                    sPeRC = "Null";
                }
                if (sPrRC != "")
                {
                    sPrRC = "" + sPrRC + "";
                }
                else
                {
                    sPrRC = "Null";
                }

                strQuery = "Update " + strCustAcc + Init_Tables.gTblStageInfo + " set Au_ReminderCancel=" + sAuRC + ",Ed_ReminderCancel=" + sEdRC + ",Pe_ReminderCancel=" + sPeRC + ",Pr_ReminderCancel=" + sPrRC + " where autoartid='" + sjobid + "' and revfinstage='fp'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                strQuery = DBProc.GetResultasString("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + strCustAcc + "', '" + sjobid + "' , '" + Session["UserID"] + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','ArticleDetailsUpdated in Jobhistory','JobHistory','Null')", Session["sConnSiteDB"].ToString());

            }
            catch (Exception ex)
            {
            }
        }
        public void UpdateLicenceCancelDetails(string sstop, string sjobid)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strQuery = "";
                if (sstop != "")
                {
                    sstop = "'" + sstop + "'";
                }
                else
                {
                    sstop = "Null";
                }
                strQuery = "Update " + strProdinfo + " set ToStop_Licencechaser=" + sstop + " where autoartid='" + sjobid + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

            }
            catch (Exception ex)
            {
            }
        }
        public void UpdateArticleDetailssingle(string VendorPgs, string CopyRights, string sjobid, string txtNoFigs, string EpubDt, string textMSP, string textTSP, string textTables, string Iss, string txtMScriptId)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strQuery = "";
                string strCustGroup = Session["sCustGroup"].ToString();

                if (textTSP != "")
                {
                    textTSP = "'" + textTSP + "'";
                }
                else
                {
                    textTSP = "Null";
                }

                if (textMSP != "")
                {
                    textMSP = "'" + textMSP + "'";
                }
                else
                {
                    textMSP = "Null";
                }
                if (txtNoFigs != "")
                {
                    txtNoFigs = "'" + txtNoFigs + "'";
                }
                else
                {
                    txtNoFigs = "Null";
                }


                if (VendorPgs != "")
                {
                    VendorPgs = "'" + VendorPgs + "'";
                }
                else
                {
                    VendorPgs = "Null";
                }

                if (textTables != "")
                {
                    textTables = "'" + textTables + "'";
                }
                else
                {
                    textTables = "Null";
                }

                if (CopyRights != "")
                {
                    CopyRights = "'" + CopyRights + "'";
                }
                else
                {
                    CopyRights = "Null";
                }

                if (txtMScriptId != "")
                {
                    txtMScriptId = "'" + txtMScriptId + "'";
                }
                else
                {
                    txtMScriptId = "Null";
                }
                if (EpubDt != "")
                {
                    EpubDt = "'" + EpubDt + "'";
                }
                else
                {
                    EpubDt = "Null";
                }
                if (Iss != "")
                {
                    Iss = "'" + Iss + "'";
                }
                else
                {
                    Iss = "Null";
                }

                strQuery = "update " + strArtcileinfo + " set NumofMSP=" + textMSP + ", NumofFigures=" + txtNoFigs + ", Iss=" + Iss + ", ActualPages=" + textTSP + ",NumofTables=" + textTables + ",VendorPgs=" + VendorPgs + ",ManuscriptID=" + txtMScriptId + ",epubdate=" + EpubDt + " where autoartid='" + sjobid + "' ";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                strQuery = "update " + strProdinfo + " set CopyRights=" + CopyRights + " where autoartid='" + sjobid + "' ";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                strQuery = DBProc.GetResultasString("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + strCustAcc + "', '" + sjobid + "' , '" + Session["UserID"] + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','ArticleDetailsUpdated in Jobhistory','JobHistory','Null')", Session["sConnSiteDB"].ToString());

            }
            catch (Exception ex)
            {
            }
        }
        public void UpdateProofDetailssingle(string proofPREMail, string proofPEEMail, string proofMailFrom, string proofEditorEMail, string proofAuthorEMail, string sjobid)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strProdinfo = strCustAcc + Init_Tables.gTblRevProdInfo;
                string strArtcileinfo = strCustAcc + Init_Tables.gTblChapterOrArticleInfo;
                string strQuery = "";

                if (proofPREMail != "")
                {
                    proofPREMail = "'" + proofPREMail + "'";
                }
                else
                {
                    proofPREMail = "Null";
                }

                if (proofPEEMail != "")
                {
                    proofPEEMail = "'" + proofPEEMail + "'";
                }
                else
                {
                    proofPEEMail = "Null";
                }
                if (proofMailFrom != "")
                {
                    proofMailFrom = "'" + proofMailFrom + "'";
                }
                else
                {
                    proofMailFrom = "Null";
                }


                if (proofEditorEMail != "")
                {
                    proofEditorEMail = "'" + proofEditorEMail + "'";
                }
                else
                {
                    proofEditorEMail = "Null";
                }

                if (proofAuthorEMail != "")
                {
                    proofAuthorEMail = "'" + proofAuthorEMail + "'";
                }
                else
                {
                    proofAuthorEMail = "Null";
                }
                strQuery = "update " + strProdinfo + " set mfrom=" + proofMailFrom + ",mto=" + proofAuthorEMail + " where AutoArtID='" + sjobid + "'";
                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                strQuery = "update j set jbm_meditor = " + proofEditorEMail + ", jbm_mpe = " + proofPEEMail + ", PrEmail = " + proofPREMail + " from JBM_Info j join  " + strArtcileinfo + " a on a.jbm_autoid = j.jbm_autoid where a.AutoArtID = '" + sjobid + "'";
                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

            }
            catch (Exception ex)
            {
            }
        }
        public void DeleteStageDetailssingle(string Stage, string sjobid)
        {
            string strCustAcc = Session["sCustAcc"].ToString();
            strStageinfo = strCustAcc + Init_Tables.gTblStageInfo;
            string strQuery = "delete from " + strStageinfo + " where autoartid='" + sjobid + "' and revfinstage='" + Stage + "'";

            try
            {
                DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                // RecordManager.CreateRecord("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + SessionHandler.sCustAcc + "', '" + strAutoartid + "' , '" + SessionHandler.sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'2', 'Deleted " + lblstage.Text + " Record in Job History','JobHistory','Null')", "");

            }
            catch (Exception ex)
            {

            }
        }
        public void UpdateStageDetailssingle(string DispatchedDate, string sjobid)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strStageinfo = strCustAcc + Init_Tables.gTblStageInfo;
                string strQuery = "";

                if (DispatchedDate != "")
                {
                    DispatchedDate = "'" + DispatchedDate + "'";
                }
                else
                {
                    DispatchedDate = "Null";
                }
                strQuery = "update " + strStageinfo + " set DispatchDate=" + DispatchedDate + " where AutoArtID='" + sjobid + "' and RevFinStage='FP'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                if (DispatchedDate != "")
                {
                    DispatchedDate = "'" + Convert.ToDateTime(DispatchedDate).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + "'";
                }
                else
                {
                    DispatchedDate = "Null";
                }



                //Label lblwf = grdHis2.Rows[e.RowIndex].FindControl("lbWflow") as Label;
                //Label lblstage = grdHis2.Rows[e.RowIndex].FindControl("lbStage") as Label;
                //Label lblCurrStage = grdHis2.Rows[e.RowIndex].FindControl("lbCurStage") as Label;

                //string lbDispatchedDate = grdHis2.DataKeys[e.RowIndex].Values["DispatchedDate"].ToString();




                //string strAutoartid = grdHis2.DataKeys[e.RowIndex].Values["jobid"].ToString();
                //string artstage = grdHis2.DataKeys[e.RowIndex].Values["ArtStage"].ToString();

                //string strArtstageTypeid = clsInit.Proc_sproofDispatched(strAutoartid, lbDispatchedDate, txtdispDate.Text, artstage, lblstage.Text);
                //string strQuery = "Update " + strStageinfo + " set  DispatchDate=" + txtdispDate.Text + " " + strArtstageTypeid + " where autoartid='" + strAutoartid + "' and revfinstage='" + lblstage.Text + "'";
                ////string strQuery = "Update " + strStageinfo + " set ReceivedDate=" + txtRecdDate.Text + ", DueDate=" + txtDueDate.Text + ", DispatchDate=" + txtdispDate.Text + " " + strArtstageTypeid + " where autoartid='" + strAutoartid + "' and revfinstage='" + lblstage.Text + "'";        //elakkiya
                //string strJAutoID = "";

                //strJAutoID = Convert.ToString(ViewState["lblAutId"]);

                //------- start for NIEHS chinese article only
                //if (strCustAccess == "EH" && strJAutoID == "EH002")
                //{
                //    strQuery += "; Update " + strProdinfo + " set AutCheckInDate=" + txtdispDate.Text + " where autoartid='" + strAutoartid + "'";
                //    RecordManager.CreateRecord("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + SessionHandler.sCustAcc + "', '" + strAutoartid + "' , '" + SessionHandler.sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','Prod info details Updated in Jobhistory','Jobhistory','Null')", "");
                //}
                ////------- End

                //RecordManager.UpdateRecord(strQuery, "");
                //if (strCustAccess == "OP" && txtdispDate.Text != "")
                //{

                //    string strStage = lblstage.Text;


                //    string cpQuery = "select j.Jbm_Autoid,j.JBM_CPJournals  from " + Init_Tables.gTblJrnlInfo + " j  where j.Jbm_Autoid='" + strJAutoID + "' and j.jbm_disabled='0'";


                //    DataTable cpDt = new DataTable();
                //    cpDt = RecordManager.GetRecord_Multiple_All(cpQuery, "CP", "");
                //    string tmpcpJournals = "";
                //    for (int i = 0; i < cpDt.Rows.Count; i++)
                //    {

                //        tmpcpJournals = cpDt.Rows[i]["JBM_CPJournals"].ToString();
                //    }


                //    if (strCustAccess == "OP" && strStage == "PapA")
                //    {


                //        RecordManager.CreateRecord("Insert into  " + Init_Tables.gTblEventInfo + " (JBM_AutoId, AutoArtID, ExternalEventID, ExternalEventStatus, ExternalEventActDate, Stage,DispatchDate) values('" + strJAutoID + "','" + strAutoartid + "','EV025','0','" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "', 'PapA'," + txtdispDate.Text + ")");

                //    }



                //    if (strCustAccess == "OP" && (strStage == "PapB" || strStage == "ePub") && tmpcpJournals.ToString().ToLower() == "true")
                //    {

                //        RecordManager.CreateRecord("Insert into  " + Init_Tables.gTblEventInfo + " (JBM_AutoId, AutoArtID, ExternalEventID, ExternalEventStatus, ExternalEventActDate, Stage,DispatchDate) values('" + strJAutoID + "','" + strAutoartid + "','EV026','0','" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "', '" + strStage + "'," + txtdispDate.Text + ")");
                //    }
                //    else if (strCustAccess == "OP" && (strStage == "PapB" || strStage == "ePub"))
                //    {


                //        RecordManager.CreateRecord("Insert into  " + Init_Tables.gTblEventInfo + " (JBM_AutoId, AutoArtID, ExternalEventID, ExternalEventStatus, ExternalEventActDate, Stage,DispatchDate) values('" + strJAutoID + "','" + strAutoartid + "','EV024','0','" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "', '" + strStage + "'," + txtdispDate.Text + ")");
                //    }



                //}


                //RecordManager.CreateRecord("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + SessionHandler.sCustAcc + "', '" + strAutoartid + "' , '" + SessionHandler.sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','Dispatch date  Updated in Jobhistory','Jobhistory','Null')", "");

            }
            catch (Exception ex)
            {
            }
        }
        public void lnkIssHold_Click(string strStage, string lblIss, string rdSearchList, string cboID, string holdStatus, string txtID)
        {
            string strJId = "";
            bool cbojrnl;
            bool boolcboIss = false;
            if (rdSearchList == "2")
            {
                boolcboIss = true;
            }

            if (cboID != "")
            {
                strJId = cboID;
                cbojrnl = true;

            }
            else
            {
                cbojrnl = false;

            }
            if (holdStatus == "Put on hold")
            {
                holdStatus = "H";
            }
            else
            {
                holdStatus = "R";

            }

            Response.Redirect("HoldArticleinfo.aspx?QS=" + holdStatus + "&Autoid=" + txtID + "&JId=" + strJId + "&Stage=" + strStage + "&CboJrnl=" + cbojrnl + "&boolIss=" + boolcboIss + "&Issue=" + lblIss.ToString() + "");

        }
        [SessionExpire]
        public ActionResult UpdateIssueDetails(string txtdispDate, string lblwf, string lblstage, string lblCurrStage, string lbljobid, string lblIss)
        {
            try
            {
                string CustDisplayId = "1";
                string strCustAccess = Session["sCustAcc"].ToString();
                string sUsrID = Session["UserID"].ToString();
                if (txtdispDate != "")
                {
                    txtdispDate = "'" + Convert.ToDateTime(txtdispDate).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + "'";
                }
                else
                {
                    txtdispDate = "Null";
                }


                string strQuery = "Update " + strCustAccess + Init_Tables.gTblIssueInfo + " set  DispatchDate=" + txtdispDate + " where jbm_autoid='" + lbljobid + "' and revfinstage='" + lblstage.Trim() + "' and iss='" + lblIss.Trim() + "'";
                if (CustDisplayId == "1")
                {
                    if (lblIss != "")
                    {
                        strQuery = "Update " + strCustAccess + Init_Tables.gTblIssueInfo + " set  DispatchDate=" + txtdispDate + " where jbm_autoid='" + lbljobid + "' and revfinstage='" + lblstage.Trim() + "' and iss='" + lblIss.Trim() + "'";
                    }
                    else
                    {
                        strQuery = "Update " + strCustAccess + Init_Tables.gTblIssueInfo + " set  DispatchDate=" + txtdispDate + " where jbm_autoid='" + lbljobid + "' and revfinstage='" + lblstage.Trim() + "'";
                    }

                }
                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());

                if (strCustAccess == "OP")
                {
                    string strStage = lblstage;
                    string strAddIss = lblIss;
                    string strJAutoID = "";
                    strJAutoID = Session["lblAutId"].ToString();

                    if (strCustAccess == "OP" && strStage == "Onl1" && txtdispDate != "")
                    {
                        //strQuery = DBProc.GetResultasString("Insert into  " + Init_Tables.gTblEventInfo + " (Jbm_autoid,iss,ExternalEventID, ExternalEventStatus,ExternalEventActDate,DispatchDate) values('" + strJAutoID + "','" + strAddIss + "','EV027','0','" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "'," + txtdispDate + ")", Session["sConnSiteDB"].ToString());
                    }
                }

                strQuery = DBProc.GetResultasString("Insert into " + strCustAccess + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + Session["sCustAcc"].ToString() + "', '" + lbljobid + "' , '" + sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'1','IssueDetails Updated in jobhistory','JobHistory','Null')", Session["sConnSiteDB"].ToString());

                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult DeleteIssueDetails(string lblstage, string lbljobid, string lblIss)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string sUsrID = "";
                string strQuery = "delete from " + strCustAcc + Init_Tables.gTblIssueInfo + " where jbm_autoid='" + lbljobid + "' and revfinstage='" + lblstage.Trim() + "' and iss='" + lblIss.Trim() + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                strQuery = DBProc.GetResultasString("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + Session["sCustAcc"].ToString() + "', '" + lbljobid + "' , '" + sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'2','DeleteRecord','JobHistory','Null')", Session["sConnSiteDB"].ToString());
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
        [SessionExpire]
        public ActionResult DeleteIssueAllocation(string lblstage, string cboID, string lblIss, string strEmpId)
        {
            try
            {
                string strCustAcc = Session["sCustAcc"].ToString();
                string strAutoartid = cboID;
                string sUsrID = "";
                string strQuery = "delete from " + strCustAcc + Init_Tables.gTblJBM_Allocation + " where autoartid='" + strAutoartid + "' and iss='" + lblIss + "' and stage='" + lblstage + "' and allocatedto='" + strEmpId + "'";

                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                strQuery = DBProc.GetResultasString("Insert into " + Init_Tables.gTblProdAccess + "(CustAcc, AutoArtID, EmpAutoID, AccTime, Process,Descript,AccPage,JBM_AutoId) values ('" + Session["sCustAcc"].ToString() + "', '" + strAutoartid + "' , '" + sUsrID + "' ,'" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yy", CultureInfo.CurrentCulture) + " " + DateTime.Now.ToLongTimeString() + "' ,'2','" + lblIss + " - " + lblstage + " - Allocation Details deleted in Jobhistory','JobHistory','Null')", Session["sConnSiteDB"].ToString());
                strQuery = DBProc.GetResultasString(strQuery, Session["sConnSiteDB"].ToString());
                return Json(new { dataComp = "Success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { dataComp = "Failed" }, JsonRequestBehavior.AllowGet);
            }

        }
        [ChildActionOnly]
        public ActionResult RenderMenu(string MenuItems)
        {
            string strNASCmd = Session["NASAccessCmd"].ToString();
            string strUserLoginName = Session["UserID"].ToString();
            string strCustGroup = Session["sCustGroup"].ToString();
            string strCustAcc = Session["sCustAcc"].ToString();
            string strNASSupportDirPath = "";
            string strCeninwDir = string.Empty;
            string strCenproDir = string.Empty;
            string strCEStyleDir = string.Empty;
            string strSGIssueDir = string.Empty;
            string strCenAppDir = string.Empty;
            string strNASInwardSupportDirPath = string.Empty;
            NasExpModel NasExpitems = new NasExpModel();

            NasExpitems.NasId = MenuItems.ToString();
            if (Session["dsartNas"] != null)
            {
                DataSet dsNas = (DataSet)Session["dsartNas"];
                string strIntrnlJID = "";
                string strSiteID = "";
                string strCustSN = "";
                string strJBMAID = "";
                string sAutoArtID = "";
                string sInternalID = "";
                string strSelVolIssNo = ""; //For NAS Explorer
                if (dsNas.Tables[0].Rows.Count > 0)
                {
                    strIntrnlJID = dsNas.Tables[0].Rows[0][31].ToString();                    
                    strSelVolIssNo = dsNas.Tables[0].Rows[0][8].ToString();    //For NAS Explorer
                    strSiteID = dsNas.Tables[0].Rows[0]["SiteID"].ToString();
                    strCustSN = dsNas.Tables[0].Rows[0]["CustSN"].ToString();
                    strJBMAID = dsNas.Tables[0].Rows[0]["AutoId"].ToString();
                    sAutoArtID = dsNas.Tables[0].Rows[0]["JobId"].ToString();
                    sInternalID = dsNas.Tables[0].Rows[0]["InternalID"].ToString();
                }
                string sArtPickupLoc = "";
                if (dsNas.Tables[1].Rows.Count > 0)
                    sArtPickupLoc = dsNas.Tables[1].Rows[0]["ArtPickup_Loc"].ToString();

                if (strSiteID == "")
                {
                    strSiteID = Session["sSiteID"].ToString();
                }

                RL.ArtDet A = new RL.ArtDet();
                A.AD.CustSN = strCustSN;
                if (sArtPickupLoc == "L0004-L0001") { A.AD.ArtPickUpSite = "L0001"; } else if (sArtPickupLoc == "L0001-L0004") { A.AD.ArtPickUpSite = "L0004"; }
                A.AD.CustAccess = strCustAcc;
                A.AD.JrnlSiteID = strSiteID;
                A.AD.VolDir = RL.clsFileIO.Proc_Extract_VolDir(strSelVolIssNo);
                A.AD.InternalJID = strIntrnlJID;
                A.AD.InternalID = sInternalID;
                A.AD.CustGroup = strCustGroup;

                string strPath = System.Web.HttpContext.Current.Server.MapPath(@"~/bin\\Smart_Config\\Smart_Config.xml");
                XmlDocument objxml = new XmlDocument();
                XmlNodeList objNodelist;
                objxml.Load(strPath);
                if (objxml.InnerXml != "Nothing")
                {
                    objNodelist = objxml.SelectNodes("//config/" + Session["sConnSiteDB"].ToString());
                    if (objNodelist.Count > 0)
                    {
                        A.AD.ConnString = objNodelist.Item(0).InnerText.ToString();
                     }
                }

                String strCeStylesheetDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F5000-CEStyleSheet", false);
                strNASSupportDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F5000-SupportProd", false);


                if (strCeStylesheetDirPath == null)
                { strCeStylesheetDirPath = ""; }
                String strInwardDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F20-01", false);//clsInit.gStrInwardFpDirPath(F20-01)
                String strCenproDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F50-01", false);
                String strCenAppDirPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F250-01", false);
                string strWorkingFolderPath = RL.clsFileIO.Proc_Get_Directory_Path(ref A, "F5000-08", false);//clsInit.gStrWorkingFolderDirPath
                strWorkingFolderPath = strWorkingFolderPath.Replace("###EmpLoginName###", strUserLoginName);

                //Label lblBKorJournal = showCenproLinks.FindControl("lblBookOrJournal") as Label;
                //string isBookJnl = lblBKorJournal.Text;

                //For NAS Explorer
                if (strCustAcc == "BK")
                {
                    strCeninwDir = strCenAppDirPath.Replace("Dispatch\\IProof\\", "");//Regex.Replace(strDispathDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");
                    strCenproDir = strCenproDirPath.Replace("ML\\", "");//Regex.Replace(strCenproDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");
                    strCenAppDir = strCenAppDirPath.Replace("Dispatch\\IProof\\", "");//Regex.Replace(strDispathDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");
                }
                else
                {
                    strCeninwDir = strInwardDirPath.Replace("FreshMss\\", ""); //Regex.Replace(strInwardDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");
                    strCenAppDir = strCenAppDirPath.Replace("Dispatch\\IProof\\", "");//Regex.Replace(strDispathDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");
                    strCenproDir = strCenproDirPath.Replace("ML\\", "");//Regex.Replace(strCenproDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1$2");

                    if (strCustAcc == "SG")
                    {
                        strSGIssueDir = "\\\\blrnas3\\cenpro\\Support\\Production\\SAGE\\Documentation\\SAGE_Issues";//'commented glyph.com by Mayavan on 27-Jun-2021
                                                                                                                     //strCEStyleDir = Regex.Replace(strDispathDirPath, "(.*?)(Vol.*?\\d{6,6})(.*)", "$1").Replace(strIntrnlJID + "\\", "") + "CE\\CE_StyleSheet\\" + strIntrnlJID + "\\";
                        if (strCeStylesheetDirPath != null)
                        {
                            strCEStyleDir = strCeStylesheetDirPath.Replace("###JID###", strIntrnlJID).Replace("\\", "[$]").Replace("#", "[@@@]");
                        }

                    }
                    else
                    {
                        if (strCeStylesheetDirPath != null)
                        {
                            strCEStyleDir = strCeStylesheetDirPath.Replace("###JID###", strIntrnlJID).Replace("\\", "[$]").Replace("#", "[@@@]");
                        }

                    }

                }

                ///////NAS Explorer
                strCeninwDir = strCeninwDir.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");
                strCenproDir = strCenproDir.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");
                strCenAppDir = strCenAppDir.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");
                strCEStyleDir = strCEStyleDir.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");
                strSGIssueDir = strSGIssueDir.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");
                strNASInwardSupportDirPath = strNASInwardSupportDirPath.Replace("\\", "[$]").Replace("#", "[@@@]").Replace("-", "[***]");

                if (strNASSupportDirPath != null)
                {
                    strNASSupportDirPath = strNASSupportDirPath.Replace("\\", "[$]").Replace("#", "[@@@]");
                }

                ////NAS Explorere New
                if (MenuItems == "NASExplorer")
                {
                    NasExpitems.data_command = strNASCmd;
                    NasExpitems.data_ceninwexp = objDS.EncryptData(strCeninwDir.Replace("smb:", ""));
                    NasExpitems.data_cenproexp = objDS.EncryptData(strCenproDir.Replace("smb:", ""));
                    NasExpitems.data_cenappexp = objDS.EncryptData(strCenAppDir.Replace("smb:", ""));
                    NasExpitems.data_workdirexp = objDS.EncryptData(strWorkingFolderPath);
                    NasExpitems.data_emp = @Session["UserID"].ToString();
                    if (strSiteID == "L0001" || A.AD.ArtPickUpSite == "L0001")
                    {
                        NasExpitems.data_nas = "withcenapp";
                    }
                    else { NasExpitems.data_nas = ""; }
                    NasExpitems.data_uniqueid = strJBMAID + "-" + sAutoArtID;
                    if (strNASSupportDirPath == null)
                    { strNASSupportDirPath = ""; }
                    NasExpitems.ToolTip = "NAS Explorer";
                }

                //// Production Support
                if (MenuItems == "ProductionSupport")
                {
                    NasExpitems.data_command = strNASCmd;
                    NasExpitems.data_CeninwExpSupport = objDS.EncryptData(strNASSupportDirPath.Replace("smb:", ""));
                    NasExpitems.data_ceninwexp = "";
                    NasExpitems.data_cenproexp = "";
                    NasExpitems.data_cenappexp = "";
                    NasExpitems.data_workdirexp = "";
                    NasExpitems.data_emp = @Session["UserID"].ToString();
                    NasExpitems.ToolTip = "NAS Explorer";
                }

                // SG Issue Explorer
                if (strCustAcc == "SG")
                {
                    if (MenuItems == "NASExplorer")
                    {
                        //    showSGIssue.Visible = true;
                        //    showSGIssueExp.Visible = true;
                        //    lblSageIssue.Visible = true;
                        NasExpitems.data_command = strNASCmd;
                        NasExpitems.data_CeninwExpSupport = objDS.EncryptData(strSGIssueDir.Replace("/", "[$]").Replace("smb:", ""));
                        NasExpitems.data_ceninwexp = "";
                        NasExpitems.data_cenproexp = "";
                        NasExpitems.data_cenappexp = "";
                        NasExpitems.data_workdirexp = "";
                        NasExpitems.data_emp = @Session["UserID"].ToString();
                        NasExpitems.ToolTip = "SAGE Issue Folder";
                    }
                }

                ////CE Stylesheet
                if (MenuItems == "CSS")
                {
                    NasExpitems.data_command = strNASCmd;
                    NasExpitems.data_CeninwExpSupport = objDS.EncryptData(strCEStyleDir.Replace("/", "[$]").Replace("smb:", ""));
                    NasExpitems.data_ceninwexp = "";
                    NasExpitems.data_cenproexp = "";
                    NasExpitems.data_cenappexp = "";
                    NasExpitems.data_workdirexp = "";
                    NasExpitems.data_emp = @Session["UserID"].ToString();
                    NasExpitems.ToolTip = "NAS Explorer";
                }


                //lblNAS.Visible = true;
                //showNASLinks.Visible = true;

                //lblSupport.Visible = true;
                //showSupportLinks.Visible = true;

                //lblCEStyle.Visible = true;
                //showCEStyleSheet.Visible = true;

                //lblCenpro.Visible = false;
                //showCenproLinks.Visible = false;

                //lblCeninw.Visible = false;
                //showCeninwLinks.Visible = false;

                /////////NAS Explorer
            }
            string query = "select RootPath from JBM_AccountTypeDesc a join JBM_RootDirectory r on r.RootID=a.InwardDir where CustAccess='" + strCustAcc + "' ";
            DataTable dt = new DataTable();
            dt = DBProc.GetResultasDataTbl(query, Session["sConnSiteDB"].ToString());


            strNASInwardSupportDirPath = dt.Rows[0]["RootPath"].ToString() + @"\Support";
            string data_CeninwExpSupport = strNASInwardSupportDirPath.Replace("smb:", "").Replace("/", "[$]");

            if (MenuItems == "InwardSupport")
            {
                NasExpitems.data_command = strNASCmd;
                NasExpitems.data_CeninwExpSupport = objDS.EncryptData(data_CeninwExpSupport);
                NasExpitems.data_emp = @Session["UserID"].ToString();
                NasExpitems.data_ceninwexp = "";
                NasExpitems.data_cenproexp = "";
                NasExpitems.data_cenappexp = "";
                NasExpitems.data_workdirexp = "";
                NasExpitems.ToolTip = "NAS Explorer";
            }

            return PartialView("NasDirectory", NasExpitems);
        }
        public ActionResult Split_Function_WFCode(string WFName)
        {
            try
            {
                string data = "";
                DataSet dscustlst = new DataSet();
                dscustlst = DBProc.GetResultasDataSet("select dbo.Split_Function_WF('" + WFName + "')", Session["sConnSiteDB"].ToString());
                if (dscustlst.Tables[0].Rows.Count > 0)
                {
                    data = dscustlst.Tables[0].Rows[0][0].ToString();
                    data = data.Replace(",", " <span style='color: #fbbf37;'>➝</span> ");
                }
                return Content(data);
            }
            catch (Exception ex)
            {
                return Json(new { dataSch = "Failed" }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}