using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarScanner
{
    public class Logger
    {
        private readonly string path;
        private readonly bool enableConsoleOutput;

        /// <summary>
        /// Logger class to write logs to a file
        /// </summary>
        /// <param name="path">File path for the logger</param>
        /// <param name="enableConsoleOutput">Whether to enable console output</param>
        public Logger(string path, bool enableConsoleOutput)
        {
            this.path = path;
            this.enableConsoleOutput = enableConsoleOutput;

            if (!File.Exists(path))
            {
                File.Create(path);
            }
        }

        /// <summary>
        /// Writes a line to the log file with a timestamp
        /// </summary>
        /// <param name="content">The content to be written on the line</param>
        public void WriteLine(string content)
        {
            using (var writer = new StreamWriter(path, true))
            {
                string formattedContent = $"[{DateTime.Now}] {content}";

                writer.WriteLine(formattedContent);

                if (enableConsoleOutput)
                {
                    Console.WriteLine(formattedContent);
                }
            }
        }

        /// <summary>
        /// Writes content to the log file without a new line at the end
        /// </summary>
        /// <param name="content">The content to be written</param>
        public void Write(string content)
        {
            using (var writer = new StreamWriter(path, true))
            {
                string formattedContent = $"[{DateTime.Now}] {content}";

                writer.WriteLine(formattedContent);

                if (enableConsoleOutput)
                {
                    Console.WriteLine(formattedContent);
                }
            }
        }
    }
}
