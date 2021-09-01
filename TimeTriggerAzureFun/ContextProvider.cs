using Microsoft.SharePoint.Client;
using System;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Logging;
using OfficeDevPnP.Core;

namespace TimeTriggerAzureFun
{
    public class ContextProvider
    {
        private ILogger _log;
        public ContextProvider(ILogger log)
        {
            _log = log;
        }

        private X509Certificate2 GetAppOnlyCertificate()
        {
            X509Certificate2 appOnlyCertificate = null;

            try
            {
                X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                certStore.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = certStore.Certificates.Find(
                    X509FindType.FindByThumbprint, Environment.GetEnvironmentVariable("AADCertificateThumbprint") ?? throw new InvalidOperationException(), false);
                _log.LogInformation($"AzureADCertificateThumbprint :  {Environment.GetEnvironmentVariable("AADCertificateThumbprint")}");
                // Get the first cert with the thumbprint
                if (certCollection.Count > 0)
                {
                    appOnlyCertificate = certCollection[0];
                }
                certStore.Close();

                if (appOnlyCertificate == null)
                {
                    throw new ArgumentNullException(nameof(appOnlyCertificate));
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex.Message, ex);
            }

            return appOnlyCertificate;
        }

        public ClientContext GetAppOnlyClientContext(string siteUrl)
        {
            if (String.IsNullOrEmpty(siteUrl))
            {
                throw new ArgumentNullException(nameof(siteUrl));
            }

            string tenantId = Environment.GetEnvironmentVariable("AADTenantId");

            X509Certificate2 certificate = GetAppOnlyCertificate();

            AuthenticationManager authManager = new AuthenticationManager();
            ClientContext context = authManager.GetAzureADAppOnlyAuthenticatedContext(
                siteUrl,
                Environment.GetEnvironmentVariable("AADClientId"),
                tenantId,
                certificate);

            return context;
        }
    }
}
