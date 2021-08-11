using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using System;

namespace Users_Permission_Comparer.Connection_Adapter
{
    public class ConnectionAdapter
    {
        public IIdentityManagementService identityManagementService { get; private set; }
        public Uri ConnUri { get; private set; }

        private TfsTeamProjectCollection Tpc;
        private Logger.Logger _Logger;

        /// <summary>
        ///  Connection constructor
        /// </summary>
        /// <param name="AzureConnection">The server URL for the connection</param>
        /// <param name="log">Logger</param>
        public ConnectionAdapter(string AzureConnection, Logger.Logger Logger)
        {
            _Logger = Logger;
            Connection(AzureConnection);
        }

        /// <summary>
        /// Connect to the Azure DevOps Server
        /// </summary>
        /// <param name="AzureConnection">The server URL for the connection</param>
        private void Connection(string AzureConnection)
        {
            try
            {
                if (!string.IsNullOrEmpty(AzureConnection))
                {
                    ConnUri = new Uri(AzureConnection);
                    Tpc = new TfsTeamProjectCollection(ConnUri);

                    // Get the identity management service
                    identityManagementService = Tpc.GetService<IIdentityManagementService>();

                    string message = string.Format("Connection was successful to {0}", AzureConnection);
                    _Logger.Info(message);
                    _Logger.Flush();
                }
                else
                {
                    _Logger.Error("Connection problem, check the connection data!");
                    _Logger.Flush();
                }
            }
            catch (Exception ex)
            {
                _Logger.Error("Problem: " + ex.Message);
                _Logger.Flush();
                Environment.Exit(-1);
            }
        }
    }
}
