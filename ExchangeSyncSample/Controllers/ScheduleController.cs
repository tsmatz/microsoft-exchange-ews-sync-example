using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using ExchangeSyncSample.Models;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json.Linq;

namespace ExchangeSyncSample.Controllers
{
    public class ScheduleController : Controller
    {
        public ActionResult Index(string code)
        {
            // redirect to Azure AD (and returning code)
            if (string.IsNullOrEmpty(code))
            {
                var clientId = ConfigurationManager.AppSettings["ClientId"];
                var redirectUri = ConfigurationManager.AppSettings["RedirectUri"];
                return Redirect(string.Format("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&response_mode=query&client_id={0}&scope=https%3a%2f%2foutlook.office.com%2fEWS.AccessAsUser.All+offline_access&redirect_uri={1}",
                    HttpUtility.UrlEncode(clientId),
                    HttpUtility.UrlEncode(redirectUri)));
            }

            // get access token form code
            HttpClient cl = new HttpClient();
            var requestBody = new List<KeyValuePair<string, string>>();
            requestBody.Add(
                new KeyValuePair<string, string>("grant_type", "authorization_code"));
            requestBody.Add(
                new KeyValuePair<string, string>("code", code));
            requestBody.Add(
                new KeyValuePair<string, string>("client_id",
                    ConfigurationManager.AppSettings["ClientId"]));
            requestBody.Add(
                new KeyValuePair<string, string>("client_secret",
                    ConfigurationManager.AppSettings["ClientSecret"]));
            requestBody.Add(
                new KeyValuePair<string, string>("scope",
                    @"https://outlook.office.com/EWS.AccessAsUser.All"));
            requestBody.Add(
                new KeyValuePair<string, string>("redirect_uri",
                    ConfigurationManager.AppSettings["RedirectUri"]));
            var resMsg1 = cl.PostAsync("https://login.microsoftonline.com/common/oauth2/v2.0/token",
                new FormUrlEncodedContent(requestBody)).Result;
            var resStr1 = resMsg1.Content.ReadAsStringAsync().Result;
            JObject json1 = JObject.Parse(resStr1);
            var tokenType = ((JValue)json1["token_type"]).ToObject<string>();
            var accessToken = ((JValue)json1["access_token"]).ToObject<string>();

            //DateTime nowDate = DateTime.Now;
            DateTime nowDate = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.Now.ToUniversalTime(), "Tokyo Standard Time");
            ViewData["Title"] =
                nowDate.Year + " 年 " + nowDate.Month + " 月";
            ViewData["OAuthToken"] = accessToken;

            return View();
        }

        // URI like http://.../Schedule/ThisMonthItems?MailAddress=...&Password=...
        public ActionResult ThisMonthItems(O365AccountModel model)
        {
            // Exchange Online に接続 (今回はデモなので、Address は決めうち !)
            string oauthToken = model.OAuthToken;
            //ExchangeVersion ver = new ExchangeVersion();
            //ver = ExchangeVersion.Exchange2010_SP1;
            //ExchangeService sv = new ExchangeService(ver, TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
            ExchangeService sv = new ExchangeService(TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
            //sv.TraceEnabled = true; // デバッグ用
            //sv.Credentials = new System.Net.NetworkCredential(emailAddress, password);
            //sv.EnableScpLookup = false;
            //sv.AutodiscoverUrl(emailAddress, AutodiscoverCallback);
            //sv.Url = new Uri(@"https://hknprd0202.outlook.com/EWS/Exchange.asmx");
            //sv.Url = new Uri(model.Url);
            sv.Credentials = new OAuthCredentials(oauthToken);
            sv.Url = new Uri(@"https://outlook.office365.com/EWS/Exchange.asmx");

            // 今月の予定 (Appointment) を取得
            //DateTime nowDate = DateTime.Now;
            DateTime nowDate = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.Now.ToUniversalTime(), "Tokyo Standard Time");
            DateTime firstDate = new DateTime(nowDate.Year, nowDate.Month, 1);
            DateTime lastDate = firstDate.AddDays(DateTime.DaysInMonth(nowDate.Year, nowDate.Month) - 1);
            CalendarView thisMonthView = new CalendarView(firstDate, lastDate);
            FindItemsResults<Appointment> appointRes = sv.FindAppointments(WellKnownFolderName.Calendar, thisMonthView);

            // 結果 (Json 値) を作成
            IList<object> resList = new List<object>();
            foreach (Appointment appointItem in appointRes.Items)
            {
                // (注意 : Json では、Date は扱えない !)
                resList.Add(new
                {
                    Subject = appointItem.Subject,
                    StartYear = appointItem.Start.Year,
                    StartMonth = appointItem.Start.Month,
                    StartDate = appointItem.Start.Day,
                    StartHour = appointItem.Start.Hour,
                    StartMinute = appointItem.Start.Minute,
                    StartSecond = appointItem.Start.Second
                });
            }

            return new JsonResult()
            {
                Data = resList,
                ContentEncoding = System.Text.Encoding.UTF8,
                ContentType = @"application/json",
                JsonRequestBehavior = JsonRequestBehavior.AllowGet
            };
        }

        //private bool AutodiscoverCallback(string url)
        //{
        //    // リダイレクトの検証をおこない、OK なら true !
        //    return true;
        //}
    }
}
