using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration.Rtd;
using System.Timers;
using Newtonsoft.Json;

namespace ExcelDnaDemo
{
    [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    [ProgId(RtdClockServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    public class RtdClockServer : ExcelRtdServer
    {
        public const string ServerProgId = "RtdClock.ClockServer";

        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        System.Timers.Timer _timer;
        List<Topic> _topics;
        private const int Throttle = 1000;  // 1sec polling max

        private void SetTimer(int rate)
        {
            rate = Math.Max(Throttle, rate);     //throttling rate refresh
            if(_timer != null)
                _timer.Dispose();

            _timer = new System.Timers.Timer(rate)
            {
                AutoReset = true
            };
            _timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            _timer.Start();
        }

        protected override bool ServerStart()
        {
            _topics = new List<Topic>();
            return true;
        }

        protected override void ServerTerminate()
        {
            _timer.Dispose();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            int.TryParse(topicInfo[0], out int rate);
            SetTimer(rate);

            _topics.Add(topic);
            return "(ConnectData)";
        }

        protected override void DisconnectData(Topic topic)
        {
            _topics.Remove(topic);
        }


        private async void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            var now = await GetNow();
            foreach (var topic in _topics)
                topic.UpdateValue(now.ToString("o"));
        }

        private async Task<DateTime> GetNow()
        { 
            using (var client = new HttpClient())
            {
                var task = await client.GetAsync("http://worldtimeapi.org/api/timezone/Europe/Paris");
                var json = await task.Content.ReadAsStringAsync();
                var clock = JsonConvert.DeserializeObject<ClockResult>(json);
                return clock.datetime;
            }
        }
         
    }
}
