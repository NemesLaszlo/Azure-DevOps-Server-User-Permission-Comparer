using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using Users_Permission_Comparer.Comparer;
using Users_Permission_Comparer.Connection_Adapter;

namespace Users_Permission_Comparer.Tool_Parser
{
    public static class ToolTaskConfigurator
    {
        // Connection informations to the Azure DevOps Server - Url with the collection
        private static string _connectionData = string.Empty;
        // Users to compare
        private static Dictionary<string, string> _users = new Dictionary<string, string>();

        /// <summary>
        /// Custom Logging Initializer
        /// </summary>
        /// <returns>Costom Logger class</returns>
        private static Logger.Logger LoggerInit()
        {
            string initializeData = string.Empty;
            ConfigurationSection diagnosticsSection = (ConfigurationSection)ConfigurationManager.GetSection("system.diagnostics");
            ConfigurationElement traceSection = diagnosticsSection.ElementInformation.Properties["trace"].Value as ConfigurationElement;
            ConfigurationElementCollection listeners = traceSection.ElementInformation.Properties["listeners"].Value as ConfigurationElementCollection;
            foreach (ConfigurationElement listener in listeners)
            {

                initializeData = listener.ElementInformation.Properties["initializeData"].Value.ToString();

            }

            // Check the path with the .log extension
            if (initializeData.Contains(".log"))
            {
                Logger.Logger log = new Logger.Logger(initializeData);
                return log;
            }
            return null;
        }

        /// <summary>
        /// Config Connection to the Azure DevOps Server and the user names to comapre the different permissions
        /// </summary>
        private static void ConfigDataInit()
        {
            // Connection Initializer
            NameValueCollection ConnectionInformations = ConfigurationManager.GetSection("Connection") as NameValueCollection;
            foreach (string info in ConnectionInformations.AllKeys)
            {
                Console.WriteLine(info + ": " + ConnectionInformations[info]);
                _connectionData = ConnectionInformations[info];
            }

            // User imformations - to compare this two users permission differences (by name)
            NameValueCollection UsersInformations = ConfigurationManager.GetSection("Users") as NameValueCollection;
            foreach (string info in UsersInformations.AllKeys)
            {
                Console.WriteLine(info + ": " + UsersInformations[info]);
                _users.Add(info, UsersInformations[info]);
            }
        }

        /// <summary>
        /// Execute the tool
        /// </summary>
        public static void ConsoleRun()
        {
            ConfigDataInit();

            Logger.Logger Logger = LoggerInit();
            // Exit and information message about the log path problem
            if (Logger == null)
            {
                Console.WriteLine("Wrong 'log path' please add the full path of the space to save the log files, " +
                    "with the desired name of this file with a '.log' extension!");
                return;
            }

            Console.WriteLine("Users Permission Comparer starts ...");
            //Connect to the server
            ConnectionAdapter ConnectionAdapter = new ConnectionAdapter(_connectionData, Logger);
            Console.WriteLine("Users Permission Comparer processing ...");
            Stopwatch watch = Stopwatch.StartNew();
            PermissionComparer permissionComparer = new PermissionComparer(ConnectionAdapter, Logger);
            permissionComparer.WriteToFileUsersPermissions(_users["First"], _users["Second"]);
            watch.Stop();
            TimeSpan timeSpan = watch.Elapsed;
            Console.WriteLine($"Total processing time: {Math.Floor(timeSpan.TotalMinutes)}:{timeSpan:ss\\.ff}");
            Console.WriteLine("Done, press any key to exit");
        }
    }
}
