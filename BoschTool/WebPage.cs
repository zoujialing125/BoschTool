using System;
using System.Net;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;

namespace BoschTool
{
    public class HttpClient : WebClient
    {
        public CookieContainer CookieContainer { get; set; }
        public WebResponse _Response = null;
        public WebRequest _Request = null;
        public bool AutoRedirect { get; set; }
        public HttpClient(CookieContainer cookieContainer)
        {
            CookieContainer = cookieContainer;
        }

        public HttpClient()
        {
            CookieContainer = new CookieContainer();
        }

        protected override WebRequest GetWebRequest(Uri uri)
        {
            WebRequest request = base.GetWebRequest(uri);
            HttpWebRequest baseRequest = request as HttpWebRequest;
            if (baseRequest is HttpWebRequest)
            { 
                (baseRequest as HttpWebRequest).KeepAlive = true;
                (baseRequest as HttpWebRequest).CookieContainer = CookieContainer;
                if (this.AutoRedirect)
                {
                    (baseRequest as HttpWebRequest).AllowAutoRedirect = true;
                }
                else
                {
                    (baseRequest as HttpWebRequest).AllowAutoRedirect = false;
                }
                
                (baseRequest as HttpWebRequest).UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142";
                (baseRequest as HttpWebRequest).Headers.Add("Accept-Language", "en-US,en;q=0.8,zh-Hans-CN;q=0.5,zh-Hans;q=0.3");
                if ((baseRequest as HttpWebRequest).Method.ToLower() == "post" && string.IsNullOrEmpty((baseRequest as HttpWebRequest).ContentType))
                {
                    (baseRequest as HttpWebRequest).ContentType = "application/x-www-form-urlencoded";
                }                
            }

            this._Request = request;
            return request;
        }

        protected override WebResponse GetWebResponse(WebRequest request)
        {
            try
            {
                this._Response = base.GetWebResponse(request);
                HttpWebResponse baseResponse = this._Response as HttpWebResponse;
            }
            catch (WebException ex)
            {
                if (this._Response == null)
                    this._Response = ex.Response;
            }
            return this._Response;
        }
    }

    public static class NameValueCollectionEx
    {
        public static NameValueCollection ToNameValueCollection<TKey, TValue>(this IDictionary<TKey, TValue> dict)
        {
            var nameValueCollection = new NameValueCollection();

            foreach (var kvp in dict)
            {
                string value = null;
                if (kvp.Value != null)
                    value = kvp.Value.ToString();

                nameValueCollection.Add(kvp.Key.ToString(), value);
            }

            return nameValueCollection;
        }
    }

    public class WebPage
    {
        public HttpClient client;
        public WebBrowser browser;
        public string host;

        private WebPage() { }

        public WebPage(CookieContainer cookieContainer, string hostAddress)
        {
            client = new HttpClient(cookieContainer);
            host = hostAddress;
            client.Encoding = Encoding.UTF8;
            
            //Setup proxy as system default
            IWebProxy proxy = WebRequest.DefaultWebProxy;
            proxy.Credentials = CredentialCache.DefaultCredentials;
            client.Proxy = proxy;
            client.AutoRedirect = false;
        }

        private static void WriteHtmlToFile(string htmlContent)
        {
            File.WriteAllText(@"C:\Users\Johney_Local\Desktop\page.html", htmlContent);
        }

        public bool GetLoginState()
        {
            string responseText = client.DownloadString(host + "/Frame/Main.aspx");
            HttpWebResponse response = client._Response as HttpWebResponse;
            if(response.StatusCode == HttpStatusCode.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public Image Captcha_IMG(ref Dictionary<string, string> arrInputs)
        {
            byte[] responseData = client.DownloadData(host + "/Login/Login.aspx");
            string response = Encoding.UTF8.GetString(responseData);
            HtmlDocument html = GetHtmlDocument(response);

            string uriCaptcha = html.GetElementById("Captcha_IMG").GetAttribute("src").Substring(6);
            arrInputs = GetInputValues(html.Forms[0], "btnLogin");
            
            Stream respStream = new MemoryStream(client.DownloadData(host + uriCaptcha));
            Image imgCaptcha = Image.FromStream(respStream);
            return imgCaptcha;
        }

        public bool LoginWebServer(Dictionary<string, string> inputsValues)
        {
            try
            {
                string postForm = GetEncodedFormData(inputsValues);
                string response = client.UploadString(host + "/Login/Login.aspx", postForm);


                response = client.DownloadString(host + "/Login/SelectSite.aspx");
                HtmlDocument html = GetHtmlDocument(response);
                
                postForm = GetEncodedFormData(GetInputValues(html.Forms[0], "btnLogin"));
                response = client.UploadString(host + "/Login/SelectSite.aspx", postForm);

                //WiteHtmlToFile(response);
                if (response.Contains(inputsValues["txtAccountName"].ToUpper()))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            
        }

        public string UploadFile(string filePath)
        {
            string result = "Upload start";
            string content = client.DownloadString(host + "/T930/BoschInvoiceImport.aspx");
            if (content.Contains("ASPxPageControl1_FileUploadControl1_UploadControl1_TextBoxT_Input"))
            {
                HtmlDocument html = GetHtmlDocument(content);

                string[] formInputFields = new string[]{
                    "__EVENTTARGET", "__EVENTARGUMENT", "__VIEWSTATE", "__VIEWSTATEGENERATOR", "__VIEWSTATEENCRYPTED"
                };

                NameValueCollection values = new NameValueCollection();
                NameValueCollection files = new NameValueCollection();
                foreach(string str in formInputFields)
                {
                    values.Add(str, html.GetElementById(str).GetAttribute("value"));
                }
                List<string> list = new List<string>();

                values.Add("ASPxPageControl1", "{&quot;activeTabIndex&quot;:0}");
                values.Add("ASPxPageControl1$FileUploadControl1$UploadControl1", "{&quot;inputCount&quot;:1}");
                files.Add("ASPxPageControl1_FileUploadControl1_UploadControl1_TextBoxT_Input", "");
                files.Add("ASPxPageControl1_FileUploadControl1_UploadControl1_TextBox0_Input", filePath);

                string boundary = "----------------------------" + DateTime.Now.Ticks.ToString("x");
                byte[] postDataA = GetMultiPartData(boundary, values, files, false);

                values.Clear();
                values.Add("ASPxPageControl1$btnImport", "Import");
                values.Add("ASPxPageControl1$hidBatchID", "");
                values.Add("ASPxPageControl1$hidShipmentId", "");
                byte[] postDataB = GetMultiPartData(boundary, values, null, true);
                byte[] postData = new byte[postDataA.Length + postDataB.Length];
                postDataA.CopyTo(postData, 0);
                postDataB.CopyTo(postData, postDataA.Length);

                content = Encoding.UTF8.GetString(postData);
                client.Headers.Add("Content-Type", "multipart/form-data; boundary=" + boundary);
                byte[] byteContent = client.UploadData(host + "/T930/BoschInvoiceImport.aspx", "POST", postData);
                content = Encoding.UTF8.GetString(byteContent);
                html = GetHtmlDocument(content);
                if (content.Contains("id=\"ASPxPageControl1_lblSuccessful\""))
                {
                    result = html.GetElementById("ASPxPageControl1_lblSuccessful").InnerText.Trim();
                }
                else
                {
                    result = "UPLOAD FAILD";
                }
            }
            return result;
        }
        public string AssignCode(string InvoiceID, string ShipmentID)
        {
            string router = "/DephiErpInterface/ShipmentInvoiceTemporary.aspx";
            Uri uri = new Uri(host + router);
            string txt = client.DownloadString(uri);
            HtmlDocument html = GetHtmlDocument(txt);
            if (html.GetElementById("ASPxPageControl1_chkNotExistsWoInvoice") != null)
            {
                html.GetElementById("ASPxPageControl1_chkNotExistsWoInvoice").SetAttribute("checked", "");
            }
            var submitValues = GetInputValues(html.Forms["form1"], "ASPxPageControl1_btnSearch");
            submitValues.Add("ASPxPageControl1$txtShipmentInvoiceId", InvoiceID);
            var FormData = GetEncodedFormData(submitValues);
            txt = client.UploadString(uri, FormData);
            html = GetHtmlDocument(txt);

            if (html.GetElementById("ASPxPageControl1_IsPageControl1_LabeltotalCount").InnerText.Trim() == "1")
            {
                string shpmtId = ShipmentIDValue(ShipmentID);
                if (!string.IsNullOrEmpty(shpmtId))
                {
                    html.GetElementById("ASPxPageControl1_grvShipmentInvoice_ctl02_chkSelect").SetAttribute("checked", "True");
                    submitValues = GetInputValues(html.Forms["form1"], "ASPxPageControl1_update_shipment_popup_btnSave");
                    submitValues.Add("ASPxPageControl1$txtShipmentInvoiceId", InvoiceID);
                    submitValues["ASPxPageControl1$update_shipment_popup$txtShipmentName"] = ShipmentID;
                    submitValues["ASPxPageControl1$update_shipment_popup$hidShipmentId"] = shpmtId;
                    FormData = GetEncodedFormData(submitValues);
                    txt = client.UploadString(uri, FormData);
                    html = GetHtmlDocument(txt);
                    if (txt.Contains("ASPxPageControl1_lblErrMessage"))
                    {
                        string result = html.GetElementById("ASPxPageControl1_lblErrMessage").InnerText;
                        return result;
                    }
                    else
                    {
                        return "Assign shipment ID failed";
                    }
                }
                else
                {
                    return "Cannot find the shipment ID";
                }
            }
            else
            {
                return "Cannot find the invoice number";
            }
            
        }

        public string ShipmentIDValue(string id)
        {
            string router = "/Pop/SelectShipment.aspx";
            Uri uri = new Uri(host + router);
            string txt = client.DownloadString(uri);
            HtmlDocument document = GetHtmlDocument(txt);
            HtmlElement table = document.GetElementById("ctl00_ContentPlaceHolder1_ASPxPageControl1_gvShipment");
            for (var i = 1; i < table.GetElementsByTagName("tr").Count; i++)
            {
                string ShipmentID = table.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[0].InnerText.Trim();
                string ShipmentIDValue = table.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[0].FirstChild.GetAttribute("value");
                if (id == ShipmentID) 
                {
                    return ShipmentIDValue;
                }
            }
            return string.Empty;
        } 

        public string GenerateCCS(string ShipmentID, string TradeType = "0110")
        {
            string route = "/Customs_Invoice/Resolutionshipment.aspx?from=M";
            Uri uri = new Uri(host + route);
            string txt = client.DownloadString(uri);
            HtmlDocument document = GetHtmlDocument(txt);
            document.GetElementById("ASPxPageControl1_txtShipmentName_I").SetAttribute("value", ShipmentID);
            document.GetElementById("ASPxPageControl1_txtCreateBeginDate_I").SetAttribute("value", "");
            var submitValues = GetInputValues(document.Forms["form1"], "ASPxPageControl1_Button1");
            submitValues.Add("ASPxPageControl1$txtCreateBeginDate$State", "{&quot;rawValue&quot;:&quot;N&quot;}");
            var FormData = GetEncodedFormData(submitValues);
            txt = client.UploadString(uri, FormData);
            document = GetHtmlDocument(txt);
            if (document.GetElementById("ASPxPageControl1_IsPageControl1_LabeltotalCount").InnerText.Trim() == "1")
            {
                document.GetElementById("ASPxPageControl1_gvReShipment_ctl02_chkSelect").SetAttribute("checked", "True");
                submitValues = GetInputValues(document.Forms["form1"], "N/A");
                submitValues.Add("ASPxPageControl1$txtCreateBeginDate$State", "{&quot;rawValue&quot;:&quot;N&quot;}");
                submitValues["__EVENTTARGET"] = "ASPxPageControl1$btnNext";
                FormData = GetEncodedFormData(submitValues);
                client.AutoRedirect = true;
                txt = client.UploadString(uri, FormData);
                client.AutoRedirect = false;

                if (txt.Contains("Trade Mode Selection"))
                {
                    document = GetHtmlDocument(txt);
                    HtmlElement table = document.GetElementById("ASPxPageControl1_gvTrade");
                    for (var i = 1; i < table.GetElementsByTagName("tr").Count; i++)
                    {
                        string tradeType = table.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[1].InnerText.Trim();
                        if (tradeType == TradeType)
                        {
                            table.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[0].FirstChild.SetAttribute("checked", "True");
                            break;
                        }
                        if (i== table.GetElementsByTagName("tr").Count - 1)
                        {

                        }
                    }
                    submitValues = GetInputValues(document.Forms["form1"], "ASPxPageControl1_btnNext");
                    FormData = GetEncodedFormData(submitValues);
                    Uri selectTradeType = client._Response.ResponseUri;
                    client.AutoRedirect = true;
                    txt = client.UploadString(selectTradeType, FormData);
                    client.AutoRedirect = false;

                    if (txt.Contains("id=\"ASPxPageControl1_btnRunTo\""))
                    {
                        document = GetHtmlDocument(txt);
                        submitValues = GetInputValues(document.Forms["form1"], "ASPxPageControl1_btnRunTo");
                        FormData = GetEncodedFormData(submitValues);
                        selectTradeType = client._Response.ResponseUri;
                        txt = client.UploadString(selectTradeType, FormData);

                        if (txt.Contains("window.alert") && txt.Contains("location.href"))
                        {
                            int positionStart = txt.LastIndexOf("window.alert('");
                            int positionEnd = txt.LastIndexOf("');location.href");
                            string stateCCS = txt.Substring(positionStart + 14, positionEnd - (positionStart + 14) + 1);
                            return stateCCS;
                        }
                        else
                        {
                            return "Create CCS failed";
                        }
                    }
                    else
                    {
                        return "Cannot select trade type";
                    }
                }
                else
                {
                    return "Cannot goto page of select trade type";
                }
            }
            else
            {
                return "Cannot find shipment ID";
            }
        }

        public string DownloadCCS(string ShipmentID, string SaveToFolderPath)
        {
            string result = "Download CCS Start";
            string route = "/CustomsDeclaration/CdRequestHeaderForm.aspx";
            Uri uri = new Uri(host + route);
            string txt = client.DownloadString(uri);
            HtmlDocument document = GetHtmlDocument(txt);
            var submitValues = GetInputValues(document.Forms["form1"], "ASPxPageControl1_CdRequestHeaderList1_btnSearch");
            submitValues.Add("ASPxPageControl1$CdRequestHeaderList1$txtCustomerAssignedId", ShipmentID);
            var FormData = GetEncodedFormData(submitValues);
            txt = client.UploadString(uri, FormData);
            document = GetHtmlDocument(txt);
            if (document.GetElementById("ASPxPageControl1_CdRequestHeaderList1_IsPageControl1_LabeltotalCount").InnerText.Trim() == "1")
            {
                document.GetElementById("ASPxPageControl1_CdRequestHeaderList1_gvCdRequestHeader_ctl02_chkSelect").SetAttribute("checked", "True");
                submitValues = GetInputValues(document.Forms["form1"], "ASPxPageControl1_CdRequestHeaderList1_btnBatchExportCCSData");
                var content = client.UploadValues(uri, submitValues.ToNameValueCollection());
                var attachment = client.ResponseHeaders["Content-Disposition"];
                if (!String.IsNullOrEmpty(attachment))
                {
                    string filename = attachment.Substring(attachment.IndexOf("=") + 1);
                    string SaveTo = SaveToFolderPath;
                    if (!Directory.Exists(SaveTo))
                        Directory.CreateDirectory(SaveTo);

                    string fileFullName = Path.Combine(SaveTo, filename);
                    if (File.Exists(fileFullName))
                        File.Delete(fileFullName);

                    var file = File.OpenWrite(fileFullName);
                    file.Write(content, 0, content.Length);
                    file.Close();
                    result = "Download CCS success";
                }
                else
                {
                    result = "Download CCS failed";
                }
            }
            else
            {
                result = "Cannot find CCS";
            }

            return result;
        }

        public HtmlDocument GetHtmlDocument(string html)
        {
            browser = new WebBrowser();
            browser.ScriptErrorsSuppressed = true; //not necessesory you can remove it
            browser.DocumentText = html;
            browser.Document.OpenNew(true);
            browser.Document.Write(html);
            browser.Refresh();
            return browser.Document;
        }
        
        public static Dictionary<string, string> GetInputValues(HtmlElement HtmlForm, string submitBtnId)
        {
            HtmlElementCollection inputElementCollection = HtmlForm.GetElementsByTagName("input");
            Dictionary<string, string> dic = new Dictionary<string, string>();
            for (var i = 0; i < inputElementCollection.Count; i++)
            {
                string id = inputElementCollection[i].Id;
                string type = inputElementCollection[i].GetAttribute("type");
                string name = inputElementCollection[i].GetAttribute("name");
                string value = inputElementCollection[i].GetAttribute("value");
                if (id == submitBtnId 
                    || (type != "submit" && type != "button" && type != "file" && type != "checkbox" && type != "radio") 
                    || (inputElementCollection[i].GetAttribute("checked") == "True")
                    || (inputElementCollection[i].GetAttribute("checked") == "checked"))
                {
                    dic.Add(name, value);
                }
            }
            return dic;
        }

        public static string GetEncodedFormData(Dictionary<string, string> InputValues)
        {
            string data = "";
            foreach (var pair in InputValues)
            {
                if (pair.Key != string.Empty)
                {
                    data += "&" + WebUtility.UrlEncode(pair.Key) + "=" + WebUtility.UrlEncode(pair.Value);
                }
            }
            data = data.Substring(1);
            return data;
        }

        private static byte[] GetMultiPartData(string boundary, NameValueCollection values, NameValueCollection files = null, bool finalSets = false)
        {
            MemoryStream requestStream = new MemoryStream();
            // The first boundary
            byte[] boundaryBytes = Encoding.UTF8.GetBytes("--" + boundary + "\r\n");
            // The last boundary
            byte[] trailer = Encoding.UTF8.GetBytes("--" + boundary + "--\r\n");
            byte[] crlf = Encoding.UTF8.GetBytes("\r\n");

            foreach (string key in values.Keys)
            {
                // Write item to stream
                byte[] formItemBytes = Encoding.UTF8.GetBytes(string.Format("Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}", key, values[key]));
                requestStream.Write(boundaryBytes, 0, boundaryBytes.Length);
                requestStream.Write(formItemBytes, 0, formItemBytes.Length);
                // Add CRLF
                requestStream.Write(crlf, 0, crlf.Length);
            }

            if (files != null)
            {
                foreach (string key in files.Keys)
                {
                    // "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    string contentType = files[key].ToLower().Contains(".xls") ? "application/vnd.ms-excel" : "application /octet-stream";
                    int bytesRead = 0;
                    byte[] buffer = new byte[2048];
                    byte[] formItemBytes = Encoding.UTF8.GetBytes(string.Format("Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n", key, Path.GetFileName(files[key]), contentType));
                    requestStream.Write(boundaryBytes, 0, boundaryBytes.Length);
                    requestStream.Write(formItemBytes, 0, formItemBytes.Length);

                    if (File.Exists(files[key]))
                    {
                        using (FileStream fileStream = new FileStream(files[key], FileMode.Open, FileAccess.Read))
                        {
                            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                            {
                                // Write file content to stream, byte by byte
                                requestStream.Write(buffer, 0, bytesRead);
                            }
                            fileStream.Close();
                        }
                        //byte[] fileByte = File.ReadAllBytes(files[key]);
                        //File.WriteAllBytes(@"C:\Users\Johney_Local\Desktop\test.xlsx", fileByte);
                        //requestStream.Write(fileByte, 0, fileByte.Length);
                    }

                    // Add CRLF
                    requestStream.Write(crlf, 0, crlf.Length);
                }
            }

            if(finalSets)
                // Write trailer and close stream
                requestStream.Write(trailer, 0, trailer.Length);
            return requestStream.ToArray();
        }
    }
}
