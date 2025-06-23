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
