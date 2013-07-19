using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json.Linq;
using System.Collections.Concurrent;

namespace ExchangeSyncSample
{
    [HubName("syncHub")]
    public class ServerHub : Hub
    {
        private static readonly ConcurrentDictionary<string, ExchangeData> exchangeClients
            = new ConcurrentDictionary<string, ExchangeData>();

        // We don't use OnDisconnected override method...
        public void Ping()
        {
            // if exchange is connected, set last update
            if (exchangeClients.ContainsKey(Context.ConnectionId))
            {
                exchangeClients[Context.ConnectionId].LastUpdate = DateTime.Now;
            }
        }

        public void ConnectExchange(string userId, string password)
        {
            try
            {
                ExchangeService sv = null;
                StreamingSubscriptionConnection subcon = null;

                // Exchange Online に接続 (今回はデモなので、Address は決めうち !)
                //ExchangeVersion ver = new ExchangeVersion();
                //ver = ExchangeVersion.Exchange2010_SP1;
                //sv = new ExchangeService(ver, TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
                sv = new ExchangeService(TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"));
                //sv.TraceEnabled = true; // デバッグ用
                sv.Credentials = new System.Net.NetworkCredential(
                    userId, password);
                sv.EnableScpLookup = false;
                sv.AutodiscoverUrl(userId, exchange_AutodiscoverCallback);

                // Streaming Notification の開始 (Windows Azure の制約から 1 分ごとに貼り直し)
                StreamingSubscription sub = sv.SubscribeToStreamingNotifications(
                    new FolderId[] { new FolderId(WellKnownFolderName.Calendar) }, EventType.Created, EventType.Modified, EventType.Deleted);
                subcon = new StreamingSubscriptionConnection(sv, 1); // only 1 minute !
                subcon.AddSubscription(sub);
                subcon.OnNotificationEvent += new StreamingSubscriptionConnection.NotificationEventDelegate(exchange_OnNotificationEvent);
                subcon.OnDisconnect += new StreamingSubscriptionConnection.SubscriptionErrorDelegate(exchange_OnDisconnect);
                subcon.OnSubscriptionError += new StreamingSubscriptionConnection.SubscriptionErrorDelegate(exchange_OnSubscriptionError);
                subcon.Open();

                // set client data (Sorry, this is not scalable !)
                CleanUpExchangeClients();
                exchangeClients.TryAdd(
                    Context.ConnectionId,
                    new ExchangeData()
                    {
                        StreamingSubscription = sub,
                        LastUpdate = DateTime.Now
                    });

                // 準備完了の送信 !
                JObject jsonObj = new JObject();
                jsonObj["MailAddress"] = new JValue(userId);
                jsonObj["Password"] = new JValue(password);
                jsonObj["ServerUrl"] = new JValue(sv.Url.ToString());
                //this.SendMessage(jsonObj.ToString());
                this.Clients.Caller.notifyEvent("Ready", jsonObj.ToString());
            }
            catch (Exception exp)
            {
                JObject jsonObj = new JObject();
                jsonObj["Message"] = new JValue(exp.Message);
                this.Clients.Caller.notifyEvent("Exception", jsonObj.ToString());
            }
        }

        public void DisconnectExchange()
        {
            if (exchangeClients.ContainsKey(Context.ConnectionId))
            {
                ExchangeData dat;
                exchangeClients.TryRemove(Context.ConnectionId, out dat);
                dat.StreamingSubscription.Unsubscribe();
            }
        }

        private void CleanUpExchangeClients()
        {
            var removeKeys = from c in exchangeClients
                             where c.Value.LastUpdate < DateTime.Now.AddMinutes(-5)
                             select c.Key;
            foreach(var key in removeKeys)
            {
                ExchangeData dat;
                exchangeClients.TryRemove(key, out dat);
                dat.StreamingSubscription.Unsubscribe();
            }
        }

        private static bool exchange_AutodiscoverCallback(string url)
        {
            // リダイレクトの検証をおこない、OK なら true !
            return true;
        }

        private void exchange_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            foreach (NotificationEvent notifyEvt in args.Events)
            {
                if ((notifyEvt is ItemEvent) &&
                    ((notifyEvt.EventType == EventType.Created) ||
                    (notifyEvt.EventType == EventType.Modified) ||
                    (notifyEvt.EventType == EventType.Deleted)))
                {
                    string messageType = string.Empty;
                    if (notifyEvt.EventType == EventType.Created)
                        messageType = "Created";
                    else if (notifyEvt.EventType == EventType.Modified)
                        messageType = "Modified";
                    else if (notifyEvt.EventType == EventType.Deleted)
                        messageType = "Deleted";
                    ItemEvent itemEvt = (ItemEvent)notifyEvt;
                    JObject jsonObj;
                    if (notifyEvt.EventType == EventType.Created)
                    {
                        Appointment ap = Appointment.Bind(args.Subscription.Service, itemEvt.ItemId);
                        jsonObj = new JObject();
                        jsonObj["Subject"] = new JValue(ap.Subject);
                        jsonObj["StartYear"] = new JValue(ap.Start.Year);
                        jsonObj["StartMonth"] = new JValue(ap.Start.Month);
                        jsonObj["StartDate"] = new JValue(ap.Start.Day);
                        jsonObj["StartHour"] = new JValue(ap.Start.Hour);
                        jsonObj["StartMinute"] = new JValue(ap.Start.Minute);
                        jsonObj["StartSecond"] = new JValue(ap.Start.Second);
                    }
                    else
                    {
                        // 注 : Modified / Deleted では、アイテムが削除されていて Bind に失敗する場合があるので、その確認をすること (ここでは、何もしない...)
                        jsonObj = new JObject();
                    }
                    if(Context.ConnectionId != null)
                        this.Clients.Client(Context.ConnectionId).notifyEvent(messageType, jsonObj.ToString());
                }
            }
        }

        private void exchange_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            // 1 分ごとに継続 (Unsbscribe のときも呼ばれるので、その際は Close になる)
            StreamingSubscriptionConnection subcon = (StreamingSubscriptionConnection)sender;
            if (exchangeClients.ContainsKey(Context.ConnectionId))
            {
                subcon.Open();
            }
            else
            {
                try { subcon.Close(); }
                catch { };
            }
        }

        private void exchange_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            // ここは、Unsbscribe のときにも呼ばれるので注意 ! (そのときは、クライアントも消えているはず)
            if(Context.ConnectionId != null)
            {
                JObject jsonObj = new JObject();
                jsonObj["Message"] = new JValue(args.Exception.Message);
                this.Clients.Client(Context.ConnectionId).notifyEvent("Exception", jsonObj.ToString());
            }
        }
    }

    public class ExchangeData
    {
        public StreamingSubscription StreamingSubscription { get; set; }
        public DateTime LastUpdate { get; set; }
    }
}