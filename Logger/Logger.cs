using System;

namespace Users_Permission_Comparer.Logger
{
    public class Logger
    {
        private TextLogTraceListener _textLogTraceListener;
        // path to take our created log files
        private string _path;
        // date pariod for the log file name
        private string DateForLogTitle = string.Empty;
        public string DatedPath { get; private set; }

        /// <summary>
        /// Constructor for the custom logger
        /// </summary>
        /// <param name="path">Path to save the log files</param>
        public Logger(string path)
        {
            _path = path;
            DateForLogTitle = DateTime.Now.ToString();
            string correctForm = DateForLogTitle.Replace("/", "-").Replace(":", "_");
            string[] str = path.Split('.');
            DatedPath = str[0] + " " + correctForm + "." + str[1];
            _textLogTraceListener = new TextLogTraceListener(DatedPath);
        }

        /// <summary>
        /// Basic log writer method, which gives the format for every line
        /// </summary>
        /// <param name="message">Message to write into the log file</param>
        /// <param name="type">Message type specification</param>
        private void WriteEntry(string message, string type)
        {
            _textLogTraceListener.WriteLine(string.Format("[{0}] [{1}] [{2}]",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                type,
                message));
        }

        /// <summary>
        /// Logger close
        /// </summary>
        public void CloseWriter()
        {
            _textLogTraceListener.Close();
        }
        /// <summary>
        /// Logger flush
        /// </summary>
        public void Flush()
        {
            _textLogTraceListener.Flush();
        }

        /// <summary>
        /// Error log message writer with Error specify type
        /// </summary>
        /// <param name="message">Message to write into the log file</param>
        public void Error(string message)
        {
            WriteEntry(message, "ERROR");
        }

        /// <summary>
        /// Error log message writer with Error specify type - Exception variant
        /// </summary>
        /// <param name="ex">Received exception</param>
        public void Error(Exception ex)
        {
            WriteEntry(ex.Message + Environment.NewLine + ex.StackTrace, "ERROR");
        }

        /// <summary>
        /// Information log writer with Info specity type for basic logging
        /// </summary>
        /// <param name="message">Message to write into the log file</param>
        public void Info(string message)
        {
            WriteEntry(message, "INFO");
        }
    }
}
