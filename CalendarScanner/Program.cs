using Google.Apis.Calendar.v3.Data;
using System.Globalization;
using System.Net.Http.Headers;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using MimeKit;
using System.Text.Json;
using System.Xml;

namespace CalendarScanner
{
    /// <summary>
    /// Wrapper for result json
    /// </summary>
    public class ResultOCR
    {
        public class OCRImage
        {
            public OCRLine[] lines { get; set; }
        }

        public class OCRLine
        {
            public int[] boundingBox { get; set; }
            public string text { get; set; }
            public OCRLine[] words { get; set; }
        }

        public string status { get; set; }
        public OCRImage recognitionResult { get; set; }
    }

    /// <summary>
    /// Wrapper for schedule format
    /// </summary>
    public class ScheduleFormat
    {
        public class ScheduleDay
        {
            public string day { get; set; }
            public int[] boundingBox { get; set; }
        }

        public ScheduleDay[] days { get; set; }
    }

    internal class Program
    {
        /// <summary>
        /// The days in a work schedule
        /// </summary>
        static readonly string[] days = new string[] { "monday", "tuesday", "wednesday", "thrusday", "friday", "saturday", "sunday" };

        /// <summary>
        /// The logger class
        /// </summary>
        static Logger logger = null;

        /// <summary>
        /// The setup info for the program
        /// </summary>
        static SetupInfo setupInfo = null;

        static void Main(string[] args)
        {
            //Initialize the setup info and logger
            setupInfo = Setup.GetInfo();
            logger = new Logger(setupInfo.LogFilePath, true);

            //Read the inbox for the first unread email and get the image attachment
            if (!ReadInbox(out MimePart imagePart))
            {
                Environment.Exit(0);
                return;
            }

            //Use Azure API to recognize the text in the image
            if (!ImageRecognition(imagePart, out ResultOCR result))
            {
                Environment.Exit(0);
                return;
            }

            //Create the events from the recognized text
            CreateEvents(result.recognitionResult);

            SendCompleteEmail();
        }

        /// <summary>
        /// Reading the box inbox for the first unread email and getting the first image attachment
        /// </summary>
        /// <param name="imagePart">The image given back</param>
        /// <returns>Succed or not</returns>
        private static bool ReadInbox(out MimePart imagePart)
        {
            imagePart = null;

            //Setup connection to the google IMAP4 server
            var mailRepository = new MailRepository("imap.gmail.com", 993, true, setupInfo.ScannerEmail, setupInfo.ScannerEmailAPIPassword);
            var unreadEmails = mailRepository.GetUnreadMails();

            //Check if we have an emails
            //If we do not we stop the program
            if (unreadEmails.Count() == 0)
            {
                logger.WriteLine("No unread emails");
                return false;
            }

            logger.WriteLine("Have unread emails");

            var workEmails = unreadEmails.Where(x => x.Subject.Contains("Work schedule")).ToHashSet();

            //Get the attachment image
            if (workEmails.Count < 1)
            {
                logger.WriteLine("No work schedule emails found in the inbox");
                return false;
            }

            var image = workEmails.FirstOrDefault().BodyParts.Where(x => x.ContentType.MediaType == "image").First();

            imagePart = (MimePart)image;

            logger.WriteLine("Got the first attachment in the email");

            //Saving the image to the bin folder
            using (var stream = File.Create(image.ContentType.Name))
            {
                imagePart.Content.DecodeTo(stream);
            }

            logger.WriteLine("Saved the attachment image from the email");

            return true;
        }

        /// <summary>
        /// Recognizing the image using the Azure Image API
        /// </summary>
        /// <param name="imagePart">The image content</param>
        /// <param name="result">The result for reading the image</param>
        /// <returns>Succed or not</returns>
        /// <exception cref="Exception">If an invalid file path happened</exception>
        private static bool ImageRecognition(MimePart imagePart, out ResultOCR result)
        {
            result = null;

            logger.WriteLine("Checking if attachment exists");

            // Get the path and filename to process from the user.
            if (imagePart == null)
            {
                logger.WriteLine("No image attachment found in the email");
                return false;
            }

            if (!File.Exists(imagePart.ContentType.Name))
            {
                logger.WriteLine("Invalid file path for image");
                throw new Exception("Invalid file path for image");
            }

            logger.WriteLine("Calling Azure Image Api");

            // Call the REST API method.
            result = MakeAnalysisRequest(imagePart.ContentType.Name).Result;
            SaveJsonFile(result);
            return true;
        }

        /// <summary>
        /// Creating events from the result of the OCR
        /// </summary>
        /// <param name="result">The results</param>
        private static void CreateEvents(ResultOCR.OCRImage result)
        {
            //Check if the result is null
            if (result == null || result.lines.Length == 0)
            {
                logger.WriteLine("No lines found in the OCR result");
                return;
            }

            logger.WriteLine("Grouping the result");

            var format = JsonSerializer.Deserialize<ScheduleFormat>(File.ReadAllText(setupInfo.ScheduleFormat));

            //Get the groups of lines for each day
            var groups = GetGroups(result, format.days);

            logger.WriteLine("Creating the events");

            //Create events for the work schedule
            CreateEvents(groups);
            logger.WriteLine("Events created");
        }

        /// <summary>
        /// Sends an email to the user that the work schedule has been made
        /// </summary>
        private static void SendCompleteEmail()
        {
            var mailSender = new MailSender("smtp.gmail.com", 465, true, setupInfo.ScannerEmail, setupInfo.ScannerEmailAPIPassword);

            var message = new MimeMessage()
            {
                Subject = setupInfo.EventName,
                Body = new TextPart("plain") { Text = setupInfo.EventLocation },
            };

            message.To.Add(new MailboxAddress(setupInfo.ConfirmationName, setupInfo.ConfirmationEmail));

            logger.WriteLine($"Sending email to {setupInfo.ConfirmationEmail}");

            mailSender.SendMessage(message);

            logger.WriteLine("Finished");
        }

        /// <summary>
        /// Creating events for the work recieved work schedule
        /// </summary>
        /// <param name="groups">The text recieved from the ocr corresponding to a string day</param>
        private static void CreateEvents(Dictionary<string, ResultOCR.OCRLine[]> groups)
        {
            var calendar = new GoogleCalendar(setupInfo.CalendarKey, setupInfo.ConfirmationEmail, setupInfo.ApplicationName);

            foreach (string day in days)
            {
                //Typically OFF is the 2nd element better to check for this cause schedule could change
                if (groups[day][1].text.Contains("OFF"))
                {
                    continue;
                }

                //Setting up the description
                StringBuilder description = new StringBuilder();

                for (int i = 1; i < groups[day].Length; i++)
                {
                    description.Append(groups[day][i].text + "\n");
                }

                //getting the date
                var date = groups[day][0].words[1].text.Split('/').Select(x => int.Parse(x)).ToArray();

                //getting the start and close time
                var selectTimes = groups[day][1].text.Split('-').Concat(groups[day][groups[day].Length - 1].text.Split('-')).ToArray();

                var times = new DateTime[] { DateTime.Parse(selectTimes[0].Replace("am", "")), DateTime.Parse(selectTimes[selectTimes.Length - 1]) };

                calendar.AddEvent(new Event()
                {
                    Summary = setupInfo.EventName,
                    Location = setupInfo.EventLocation,
                    Description = description.ToString(),

                    Start = new EventDateTime()
                    {
                        DateTimeDateTimeOffset = new DateTime(DateTime.Today.Year, date[0], date[1], times[0].Hour, times[0].Minute, 0),
                        TimeZone = setupInfo.EventTimeZone
                    },
                    End = new EventDateTime()
                    {
                        DateTimeDateTimeOffset = new DateTime(DateTime.Today.Year, date[0], date[1], times[1].Hour, times[1].Minute, 0),
                        TimeZone = setupInfo.EventTimeZone
                    }
                });

                logger.WriteLine($"Created an event for {day}");
            }
        }

        /// <summary>
        /// Getting the result from the ocr and returning each lines group
        /// </summary>
        /// <param name="result">The result from json parsing</param>
        /// <param name="format">How the text in the image should be formated</param>
        /// <returns></returns>
        private static Dictionary<string, ResultOCR.OCRLine[]> GetGroups(ResultOCR.OCRImage result, ScheduleFormat.ScheduleDay[] format)
        {
            var groups = new Dictionary<string, ResultOCR.OCRLine[]>();
            for (int i = 0; i < days.Length; i++)
            {
                groups.Add(days[i], null);
            }

            for (int i = 0; i < format.Length; i++)
            {
                groups[days[i]] = result.lines.Where(x => x.boundingBox[0] >= format[i].boundingBox[0] && x.boundingBox[2] <= format[i].boundingBox[0] + format[i].boundingBox[2]).ToArray();
            }

            return groups;
        }

        /// <summary>
        /// Gets the analysis of the specified image file by using
        /// the Computer Vision REST API.
        /// </summary>
        /// <param name="imageFilePath">The image file to analyze.</param>
        static async Task<ResultOCR> MakeAnalysisRequest(string imageFilePath)
        {
            try
            {
                HttpClient client = new HttpClient();

                // Request headers.
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", setupInfo.AzureKey);

                string requestParameters = "mode=Printed";

                // Assemble the URI for the REST API method.
                string uri = setupInfo.AzureEndpoint + "?" + requestParameters;

                HttpResponseMessage response;

                // Read the contents of the specified local image
                // into a byte array.
                byte[] byteData = GetImageAsByteArray(imageFilePath);

                // Add the byte array as an octet stream to the request body.
                using (ByteArrayContent content = new ByteArrayContent(byteData))
                {
                    // This example uses the "application/octet-stream" content type.
                    // The other content types you can use are "application/json"
                    // and "multipart/form-data".
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    // Asynchronously call the REST API method.
                    // It will return Accepted (HTTP 202) immediately and include an "Operation-Location" header.
                    //  Client side should further query the operation status using the URL specified in this header.The operation ID will expire in 48 hours.
                    response = await client.PostAsync(uri, content);
                }

                logger.WriteLine("Got info from the api");

                if (!response.Headers.TryGetValues("Operation-Location", out var locations))
                {
                    logger.WriteLine("Something went terribly wrong");
                    throw new Exception("Something went terribly wrong");
                }

                logger.WriteLine("Nothing went terribly wrong");

                var location = locations.First();

                logger.WriteLine($"Image being processed {location}");

                //Keeping looping till azure image is done processing the image
                ResultOCR imageInfo = new ResultOCR();
                string contentString = "";
                while (imageInfo.status != "Succeeded")
                {
                    response = await client.GetAsync(location);
                    contentString = await response.Content.ReadAsStringAsync();
                    imageInfo = JsonSerializer.Deserialize<ResultOCR>(contentString);

                    logger.WriteLine($"Image Processing Status: {imageInfo.status}");

                    //Sleep on the calls so that the api does not kick us out
                    Thread.Sleep(100);
                }

                logger.WriteLine("Saving json file to jsonOutPut.json");

                //Save the result
                SaveJsonFile(imageInfo);

                return imageInfo;
            }
            catch (Exception e)
            {
                logger.WriteLine(e.Message);
            }

            return null;
        }

        /// <summary>
        /// Returns the contents of the specified file as a byte array.
        /// </summary>
        /// <param name="imageFilePath">The image file to read.</param>
        /// <returns>The byte array of the image data.</returns>
        static byte[] GetImageAsByteArray(string imageFilePath)
        {
            // Open a read-only file stream for the specified file.
            using (FileStream fileStream = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read))
            {
                // Read the file's contents into a byte array.
                BinaryReader binaryReader = new BinaryReader(fileStream);
                return binaryReader.ReadBytes((int)fileStream.Length);
            }
        }

        /// <summary>
        /// Save the result from the azure image api
        /// </summary>
        /// <param name="response">The result from the api</param>
        static void SaveJsonFile(ResultOCR response)
        {
            string temp = JsonSerializer.Serialize(response, new JsonSerializerOptions() { WriteIndented = true });
            File.WriteAllText("jsonOutPut.json", temp);
        }
    }
}