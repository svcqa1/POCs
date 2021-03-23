using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Web;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Graphics;
using System.Drawing;
using Syncfusion.Pdf.Interactive;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Pdf.IO;
using System.Configuration;
using Syncfusion.Pdf.Grid;
using System.IO;

namespace EffectiveDate
{
    class Program
    {
        private static string docLinkUrl;
        private static string docFileName;
        private static string spHeaderLine1 = "";
        private static string spHeaderLine2 = "";
        private static string spHeaderLine3 = "";
        private static string spWaterMark = "";
        private static string spFooterLine1 = "";
        private static string spFooterLine2 = "";
        private static string spFooterLine3 = "";
        private static string spExpiryDays = "";

        static void Main(string[] args)
        {
            string siteUrl = "https://celitotechcom.sharepoint.com/sites/QualityDocsTestEnv";
            string strPdfConfigListName = "pdfListConfig";
            string strPdfConfigListId = "1";
            string strEffectiveListName = "All Effective Documents";
            string strUserID = "svcqa1@celitotech.com";
            string strPassword = "Welcome1$";
            string strDcrDocId = "DCR132";
            getHeaderWatermarkFooter(siteUrl, strUserID, strPassword, strPdfConfigListName, strPdfConfigListId);
            updateEffectiveDate(siteUrl,strUserID,strPassword, strEffectiveListName, strDcrDocId);            
        }

        public static void getHeaderWatermarkFooter(string strSiteUrl, string strUserId, string strPassword, string strPdfConfigListName,string strPdfConfigListId)
        {
            try
            {
                using (var context = new ClientContext(strSiteUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in strPassword.ToCharArray()) passWord.AppendChar(c);
                    context.Credentials = new SharePointOnlineCredentials(strUserId, passWord);
                    // Gets list object using the list Url  
                    List oList = context.Web.Lists.GetByTitle(strPdfConfigListName);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                        "<Value Type='Text'>" + strPdfConfigListId +
                        "</Value></Eq></Where></Query></View>";
                    ListItemCollection collListItem = oList.GetItems(camlQuery);
                    context.Load(collListItem);
                    context.ExecuteQuery();
                    foreach (ListItem item in collListItem)
                    {
                        spHeaderLine1 = Convert.ToString(item["headerLine1"]);
                        spHeaderLine2 = Convert.ToString(item["headerLine2"]);
                        spHeaderLine3 = Convert.ToString(item["headerLine3"]);
                        spFooterLine1 = Convert.ToString(item["footerLine1"]);
                        spFooterLine2 = Convert.ToString(item["footerLine2"]);
                        spFooterLine3 = Convert.ToString(item["footerLine3"]);
                        spWaterMark = Convert.ToString(item["waterMark"]);
                        spExpiryDays = Convert.ToString(item["ExpiryDays"]);
                    }                                     

                }
            }
            catch (Exception ex)
            {
            }

        }
         public static void updateEffectiveDate(string strSiteUrl,string strUserId, string strPassword, string strEffectiveListName,string strDcrDocId)
         {     
            try
            {
                using (var context = new ClientContext(strSiteUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in strPassword.ToCharArray()) passWord.AppendChar(c);
                    context.Credentials = new SharePointOnlineCredentials(strUserId, passWord);
                    // Gets list object using the list Url  
                    List oList = context.Web.Lists.GetByTitle(strEffectiveListName);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='CE_DCRDocID'/>" +
                        "<Value Type='Lookup'>" + strDcrDocId +
                        "</Value></Eq></Where></Query></View>";
                    ListItemCollection collListItem = oList.GetItems(camlQuery);
                    context.Load(collListItem);
                    context.ExecuteQuery();
                    foreach (ListItem item in collListItem)
                    {
                        var filePath = item["FileRef"];
                        docFileName = (string)item["FileLeafRef"];
                       
                        
                        Microsoft.SharePoint.Client.File file = item.File;
                        docLinkUrl = new Uri(context.Url).GetLeftPart(UriPartial.Authority) + filePath;
                        if (file != null)
                        {
                            //Loading Uploaded file
                            context.Load(file);
                            context.ExecuteQuery();
                            string dskFilePath = System.IO.Path.Combine(@"C:\Srini\Celitotech\Dev\PDF files\", file.Name);
                            using (System.IO.FileStream Local_stream = System.IO.File.Open(dskFilePath, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite))
                            {
                                var fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                                var Sp_Stream = fileInformation.Stream;
                                Sp_Stream.CopyTo(Local_stream);
                            }
                         
                                PdfLoadedDocument loadedDocument = new PdfLoadedDocument(dskFilePath);
                                PdfDocument document = new PdfDocument();
                                document.ImportPageRange(loadedDocument, 0, loadedDocument.Pages.Count - 1);
                                document.Template.Top = AddHeader(document, spHeaderLine1, spHeaderLine2, spHeaderLine3, spWaterMark, spExpiryDays);
                                document.Template.Bottom = AddFooter(document,spFooterLine1,spFooterLine2,spFooterLine3);

                                for (int i = 0; i < document.Pages.Count; i++)
                                {
                                    PdfPageBase loadedPage = document.Pages[i];
                                    PdfGraphics graphics = loadedPage.Graphics;
                                    PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 45);
                                    
                                    //Add watermark text
                                    PdfGraphicsState state = graphics.Save();
                                    graphics.SetTransparency(0.25f);
                                    graphics.RotateTransform(-40);
                                    string text = spWaterMark;
                                    SizeF size = font.MeasureString(text);
                                    graphics.DrawString(spWaterMark, font, PdfPens.Gray, PdfBrushes.Gray, new PointF(-150, 450));
                                }

                            DateTime currentTime = DateTime.Now;
                            Double dbleExpiryDays = Convert.ToDouble(spExpiryDays);
                            DateTime addDaysToCurrentTime = currentTime.AddDays(dbleExpiryDays);
                            PdfJavaScriptAction scriptAction = new PdfJavaScriptAction("function Expire(){ var currentDate = new Date();  var expireDate = new Date(2021, 3, 8);      if (currentDate < expireDate) {  app.alert(\"This Document has Expired.  You need a new one.\"); this.closeDoc();   }  } Expire(); ");
                            document.Actions.AfterOpen = scriptAction;

                            document.Save(docFileName);
                            document.Close(true);
                            loadedDocument.Close(true);
                        }
                    }              
                }
            }
            catch (Exception ex)
            {
            }            
        }
              
        public static PdfPageTemplateElement AddHeader(PdfDocument doc,string headerLine1,string headerLine2,string headerLine3,string waterMark, string expiryDays)
        {    
            RectangleF rect = new RectangleF(0, 0, doc.Pages[0].GetClientSize().Width, 55);
            PdfPageTemplateElement header = new PdfPageTemplateElement(rect);
            PdfGrid pdfGrid = new PdfGrid();

            //Add three columns
            pdfGrid.Columns.Add(3);
            PdfBrush brush = new PdfSolidBrush(Color.Black);
            PdfStringFormat format = new PdfStringFormat();
            format.Alignment = PdfTextAlignment.Center;
            format.LineAlignment = PdfVerticalAlignment.Middle;
            string dtNow = DateTime.Now.ToShortDateString();
            string strSop = "SOP-2023";
            //Add rows
            PdfGridRow pdfGridRow = pdfGrid.Rows.Add();
            pdfGridRow.Cells[0].Value = headerLine1 + ":" + strSop;
            pdfGridRow.Cells[1].Value = headerLine1;
            pdfGridRow.Cells[2].Value = headerLine3 + ":" + dtNow;
                      
            //Create the font for setting the style
            PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 10, PdfFontStyle.Bold);

            for (int i = 0; i < pdfGrid.Rows.Count; i++)
            {
                PdfGridRow row = pdfGrid.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {
                    row.Cells[j].Style.Font = font;
                    row.Cells[j].StringFormat = format;
                    row.Cells[j].Style.Borders.All = PdfPens.Transparent;
                }
            }
            pdfGrid.Draw(header.Graphics);
            return header;
        }

        public static PdfPageTemplateElement AddFooter(PdfDocument doc,string strFooterLine1,string strFooterLine2,string strFooterLine3)
        {
            RectangleF rect = new RectangleF(0, 0, doc.Pages[0].GetClientSize().Width, 50);

            //Create a page template
            PdfPageTemplateElement footer = new PdfPageTemplateElement(rect);

            //Create a PdfGrid
            PdfGrid pdfGrid = new PdfGrid();

            //Add three columns
            pdfGrid.Columns.Add(1);

            //create and customize the string formats
            PdfStringFormat format = new PdfStringFormat();
            format.Alignment = PdfTextAlignment.Center;
            format.LineAlignment = PdfVerticalAlignment.Middle;

            //Create the font for setting the style
            PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 8, PdfFontStyle.Bold | PdfFontStyle.Italic);

            //Add rows
            PdfGridRow pdfGridRow = pdfGrid.Rows.Add();
            pdfGridRow.Cells[0].Value = spFooterLine1;
            pdfGridRow.Cells[0].Style.Font = font;
            pdfGridRow.Cells[0].Style.TextBrush = PdfBrushes.Red;

            //Add rows
            PdfGridRow pdfGridRow1 = pdfGrid.Rows.Add();
            pdfGridRow1.Cells[0].Value = spFooterLine2;

            //Add rows
            PdfGridRow pdfGridRow2 = pdfGrid.Rows.Add();
            pdfGridRow2.Cells[0].Value = spFooterLine3;

            for (int i = 0; i < pdfGrid.Rows.Count; i++)
            {
                PdfGridRow row = pdfGrid.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {
                    row.Cells[j].StringFormat = format;
                    row.Cells[j].Style.Borders.All = PdfPens.Transparent;
                }
            }

            //Draw grid to the page of PDF document
            pdfGrid.Draw(footer.Graphics);

            return footer;
        }
    }
}
