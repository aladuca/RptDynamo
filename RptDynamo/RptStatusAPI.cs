using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Diagnostics;


namespace RptDynamo
{
    // Data Model Definition
    public class JobStatus
    {
        public Guid ID { get; set; }
        public Status status { get; set; }
        public DateTime start { get; set; }
        public DateTime end { get; set; }
        public String filename { get; set; }
        public String requestor { get; set; }
        public String worker { get; set; }
        public Int32 processID { get; set; }
    }
    public enum Status
    {
        Queued = 0,
        Processing = 1,
        Completed = 2,
        Failed = 3
    }

    class RptStatusAPI
    {
        private JobStatus status;

        public RptStatusAPI() { status = new JobStatus(); }

        // Get Status Data
        public async Task processing(Guid ID)
        {
            using (var client = new HttpClient())
            {
                //Setup Variables for HTTP API Request
                client.BaseAddress = new Uri("http://localhost:45732/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage response;

                //Get Status Data
                response = await client.GetAsync("api/status/" + ID);
                if (response.IsSuccessStatusCode)
                {
                    status = await response.Content.ReadAsAsync<JobStatus>();
                }

                status.start = DateTime.Now;
                status.status = Status.Processing;
                status.worker = GetLocalhostFqdn().ToUpper();
                status.processID = Process.GetCurrentProcess().Id;

                //Post Status Processing
                response = await client.PostAsJsonAsync("api/status", status);
            }
        }
        public async Task completed()
        {
            using (var client = new HttpClient())
            {
                //Setup Variables for HTTP API Request
                client.BaseAddress = new Uri("http://localhost:45732/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage response;

                status.end = DateTime.Now;
                status.status = Status.Completed;
                status.processID = 0;

                response = await client.PostAsJsonAsync("api/status", status);
            }
        }
        public async Task failed()
        {
            using (var client = new HttpClient())
            {
                //Setup Variables for HTTP API Request
                client.BaseAddress = new Uri("http://localhost:45732/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                status.end = DateTime.Now;
                status.status = Status.Failed;
                status.processID = 0;

                HttpResponseMessage response = await client.PostAsJsonAsync("api/status", status);
            }
        }


        // Helper for FQDN of host: http://stackoverflow.com/questions/804700/how-to-find-fqdn-of-local-machine-in-c-net
        private static string GetLocalhostFqdn()
        {
            var ipProperties = IPGlobalProperties.GetIPGlobalProperties();
            return string.Format("{0}.{1}", ipProperties.HostName, ipProperties.DomainName);
        }

    }
}
