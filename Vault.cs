using System;
using Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;

namespace VaultFileExtraction
{
    public class Vault
    {
        public Connection Connect(
            string serverName,
            string vaultName)
        {
            var result =
                Autodesk.DataManagement.Client.Framework.Vault.Library
                .ConnectionManager.LogIn(
                    serverName,
                    vaultName,
                    string.Empty,
                    string.Empty,
                    AuthenticationFlags.WindowsAuthentication,
                    null);

            if (result == null || !result.Success)
            {
                throw new InvalidOperationException(
                    "Unable to connect to Autodesk Vault.");
            }

            return result.Connection;
        }

        public void Disconnect()
        {
            Autodesk.DataManagement.Client.Framework.Vault.Library
                .ConnectionManager.CloseAllConnections();
        }
    }
}
