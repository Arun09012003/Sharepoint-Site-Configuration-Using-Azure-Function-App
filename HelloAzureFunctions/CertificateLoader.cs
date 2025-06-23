using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.ConstrainedExecution;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace HelloAzureFunctions
{
    internal class CertificateLoader
    {
        private readonly ILogger<CertificateLoader> log;
        //private static readonly string tenantId = "a2e4046e-33db-4a6b-b46c-17f953f9ba96";
        //private static readonly string clientId = "293f06c2-8425-47de-b12d-1c78cd3f6a0d";
        //private static readonly string certificateBase64 = "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";
        //private static readonly string certificatePassword = "P@55w0rd";

        private static readonly string tenantId = Environment.GetEnvironmentVariable("TenantID");
        private static readonly string clientId = Environment.GetEnvironmentVariable("ClientID");
        private static readonly string certificateBase64 = Environment.GetEnvironmentVariable("certificateBase64");
        private static readonly string certificatePassword = Environment.GetEnvironmentVariable("certificatePassword");
        public CertificateLoader(ILogger<CertificateLoader> logger)
        {
            log = logger;
        }
        public static async Task<string> GetAccessTokenAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
                {
                    throw new Exception("TenantId or ClientId is missing in environment settings.");
                }

                 if (string.IsNullOrEmpty(certificateBase64))
                {
                    throw new Exception("Base64 certificate not found in environment settings.");
                }

                byte[] certBytes;
                try
                {
                    certBytes = Convert.FromBase64String(certificateBase64);
                    Console.WriteLine("Certificate Converted in to Base64");

                }
                catch (FormatException fe)
                {
                    throw new Exception("Invalid Base64 certificate format.", fe);
                }

                X509Certificate2 cert;
                try
                {
                    cert = new X509Certificate2(
                    certBytes,
                    certificatePassword,
                    X509KeyStorageFlags.MachineKeySet |
                    X509KeyStorageFlags.Exportable |
                    X509KeyStorageFlags.PersistKeySet
                    );
                }
                catch (Exception ce)
                {
                    throw new Exception("Failed to load certificate.", ce);
                }

                var app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .WithCertificate(cert)
                .Build();

                AuthenticationResult result;
                try
                {
                    result = await app.AcquireTokenForClient(new[] { "https://dgneaseteq.sharepoint.com/.default" }).ExecuteAsync();
                }
                catch (MsalServiceException mse)
                {
                    throw new Exception($"MSAL service error: {mse.Message}", mse);
                }
                catch (MsalClientException mce)
                {
                    throw new Exception($"MSAL client error: {mce.Message}", mce);
                }

                if (string.IsNullOrEmpty(result.AccessToken))
                {
                    throw new Exception("Access token is null or empty after acquisition.");
                }
                Console.WriteLine("Token Senging..");
                return result.AccessToken;

            }
            catch (Exception ex)
            {
                throw new Exception("Error in GetAccessTokenAsync: " + ex.Message, ex);
            }
        }


        public class TokenCredentialAuthProvider : IAuthenticationProvider
        {
            private readonly TokenCredential _tokenCredential;
            private readonly string[] _scopes;

            public TokenCredentialAuthProvider(ClientCertificateCredential tokenCredential, string[] scopes)
            {
                _tokenCredential = tokenCredential;
                _scopes = scopes;
            }
            public async Task AuthenticateRequestAsync(HttpRequestMessage request)
            {
                var token = await _tokenCredential.GetTokenAsync(new Azure.Core.TokenRequestContext(_scopes), default);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            }

            public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage, System.Threading.CancellationToken cancellationToken)
            {
                await AuthenticateRequestAsync(requestMessage);
            }

        }
        public static GraphServiceClient GetGraphClient()
        {
            byte[] certBytes;
            try
            {
                certBytes = Convert.FromBase64String(certificateBase64);
                Console.WriteLine("Certificate Converted in to Base64");

            }
            catch (FormatException fe)
            {
                throw new Exception("Invalid Base64 certificate format.", fe);
            }
            // Load certificate
            X509Certificate2 cert;
            try
            {
                cert = new X509Certificate2(
                certBytes,
                certificatePassword,
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.PersistKeySet
                );
            }
            catch (Exception ce)
            {
                throw new Exception("Failed to load certificate.", ce);
            }

            // Create ClientCertificateCredential
            var clientCertCredential = new ClientCertificateCredential(
                tenantId,
                clientId,
                cert
            );

            var authProvider = new TokenCredentialAuthProvider(clientCertCredential, new[] { "https://graph.microsoft.com/.default" });

            // Create GraphServiceClient with this credential and scopes
            var graphClient = new GraphServiceClient(authProvider);

            return graphClient;
        }

    }
}
