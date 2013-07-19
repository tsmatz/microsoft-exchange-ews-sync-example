using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExchangeSyncSample.Models;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSyncSample.Controllers
{
    public class ScheduleController : Controller
    {
        public ActionResult Index()
        {
            //DateTime nowDate = DateTime.Now;
            DateTime nowDate = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.Now.ToUniversalTime(), "Tokyo Standard Time");
            ViewData["Title"] =
                nowDate.Year + " 年 " + nowDate.Month + " 月";

            return View();
        }

        // URI like http://.../Schedule/ThisMonthItems?MailAddress=...&Password=...
        public ActionResult ThisMonthItems(O365AccountModel model)
        {
            // Exchange Online に接続 (今回はデモなので、Address は決めうち !)
            string emailAddress = model.MailAddress;
            string password = model.Password;
            //ExchangeVersion ver = new ExchangeVersion();
            //ver = ExchangeVersion.Exchange2010_SP1;
            //ExchangeService sv = new ExchangeService(ver, TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
            ExchangeService sv = new ExchangeService(TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
            //sv.TraceEnabled = true; // デバッグ用
            sv.Credentials = new System.Net.NetworkCredential(emailAddress, password);
            //sv.EnableScpLookup = false;
            //sv.AutodiscoverUrl(emailAddress, AutodiscoverCallback);
            //sv.Url = new Uri(@"https://hknprd0202.outlook.com/EWS/Exchange.asmx");
            sv.Url = new Uri(model.Url);

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
