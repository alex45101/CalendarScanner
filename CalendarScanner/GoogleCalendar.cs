using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarScanner
{
    public class GoogleCalendar
    {
        private string JsonCredentials;
        private string CalendarId;

        private ServiceAccountCredential Credential;
        private CalendarService Service;

        /// <summary>
        /// Setting up connection to google calendar
        /// </summary>
        /// <param name="jsonCredentials">Google service account credentials</param>
        /// <param name="calendarId">The unique calendar id</param>
        public GoogleCalendar(string jsonCredentials, string calendarId, string applicationName)
        {
            JsonCredentials = jsonCredentials;
            CalendarId = calendarId;

            SetupConnection(applicationName);
        }

        /// <summary>
        /// Setting up the credentials and calendar 
        /// </summary>
        private void SetupConnection(string applicationName)
        {
            using (var stream = new FileStream(JsonCredentials, FileMode.Open, FileAccess.Read))
            {
                var confg = Google.Apis.Json.NewtonsoftJsonSerializer.Instance.Deserialize<JsonCredentialParameters>(stream);

                Credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(confg.ClientEmail)
                {
                    Scopes = new string[] { CalendarService.Scope.Calendar }
                }.FromPrivateKey(confg.PrivateKey));
            }

            Service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = applicationName,
            });
        }

        /// <summary>
        /// Adding an event to the calendar
        /// </summary>
        /// <param name="evnt">The event to add</param>
        /// <returns>true if we successful at adding the event</returns>
        public bool AddEvent(Event evnt)
        {
            var InsertRequest = Service.Events.Insert(evnt, CalendarId);

            try
            {
                InsertRequest.Execute();
            }
            catch (Exception e)
            {
                try
                {
                    Service.Events.Update(evnt, CalendarId, evnt.Id).Execute();
                    return true;
                }
                catch (Exception ee)
                {
                    throw new Exception(ee.Message);
                }

                throw new Exception(e.Message);
            }

            return false;
        }
    }
}
