using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DotNetEnv;

namespace CalendarScanner
{
    public class SetupInfo
    {
        public string AzureEndpoint { get; set; }
        public string AzureKey { get; set; }
        public string ScannerEmail { get; set; }
        public string ScannerEmailAPIPassword { get; set; }
        public string CalendarKey { get; set; }
        public string ApplicationName { get; set; }
        public bool SendConfirmation { get; set; }
        public string ConfirmationName { get; set; }
        public string ConfirmationSubject { get; set; }
        public string ConfirmationBody { get; set; }
        public string ConfirmationEmail { get; set; }
        public string EventName { get; set; }
        public string EventLocation { get; set; }
        public string EventTimeZone { get; set; }
        public string ScheduleFormat { get; set; }
        public string LogFilePath { get; set; }

        public SetupInfo(string azureEndpoint, string azureKey, string scannerEmail, string scannerEmailAPIPassword, string calendarKey, string applicationName, bool sendConfirmation, string confirmationName, string confirmationEmail, string confirmationSubject, string confirmationBody, string eventName, string eventLocation, string eventTimeZone, string scheduleFormat, string logFilePath)
        {
            AzureEndpoint = azureEndpoint;
            AzureKey = azureKey;
            ScannerEmail = scannerEmail;
            ScannerEmailAPIPassword = scannerEmailAPIPassword;
            CalendarKey = calendarKey;
            ApplicationName = applicationName;
            SendConfirmation = sendConfirmation;
            ConfirmationName = confirmationName;
            ConfirmationEmail = confirmationEmail;
            ConfirmationSubject = confirmationSubject;
            ConfirmationBody = confirmationBody;
            EventName = eventName;
            EventLocation = eventLocation;
            EventTimeZone = eventTimeZone;
            ScheduleFormat = scheduleFormat;
            LogFilePath = logFilePath;
        }
    }

    public static class Setup
    {
        public static SetupInfo GetInfo()
        {
            Env.Load();

            return new SetupInfo(
                azureEndpoint: Environment.GetEnvironmentVariable("AZURE_ENDPOINT"),
                azureKey: Environment.GetEnvironmentVariable("AZURE_KEY"),
                scannerEmail: Environment.GetEnvironmentVariable("EMAIL_SCANNER_ADDRESS"),
                scannerEmailAPIPassword: Environment.GetEnvironmentVariable("EMAIL_APP_PASSWORD"),
                calendarKey: Environment.GetEnvironmentVariable("GOOGLE_CREDENTIALS_JSON"),
                applicationName: Environment.GetEnvironmentVariable("APPLICATION_NAME"),
                sendConfirmation: Environment.GetEnvironmentVariable("SEND_CONFIRMATION") == "True",
                confirmationName: Environment.GetEnvironmentVariable("NAME_CONFIRMATION"),
                confirmationEmail: Environment.GetEnvironmentVariable("EMAIL_CONFIRMATION"),
                confirmationSubject: Environment.GetEnvironmentVariable("SUBJECT_CONFIRMATION"),
                confirmationBody: Environment.GetEnvironmentVariable("BODY_CONFIRMATION"),
                eventName: Environment.GetEnvironmentVariable("EVENT_NAME"),
                eventLocation: Environment.GetEnvironmentVariable("EVENT_LOCATION"),
                eventTimeZone: Environment.GetEnvironmentVariable("EVENT_TIMEZONE"),
                scheduleFormat: Environment.GetEnvironmentVariable("SCHEDULE_FORMAT"),
                logFilePath: Environment.GetEnvironmentVariable("LOG_FILE")
            );
        }
    }
}
