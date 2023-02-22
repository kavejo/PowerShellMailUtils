using Microsoft.Identity.Client;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Management.Automation;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace SyntheticTransactionsForExchange.AzureADAuthentication
{
    [Cmdlet("Get", "ApplicationAccesToken")]
    [OutputType(typeof(String))]
    [CmdletBinding()]

    public class GetApplicationAccesToken : Cmdlet
    {
        /// <summary>
        /// <para type="description">The ClientId set on the Azure AD Application.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The ClientId set on the Azure AD Application")]
        public String ClientId = String.Empty;

        /// <summary>
        /// <para type="description">The TenantId shown for the Azure AD Application.</para>
        /// <para type="description">If unset this would default to "https://login.microsoftonline.com/common/"</para> 
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The TenantId shown for the Azure AD Application")]
        public String TenantId = String.Empty;

        /// <summary>
        /// <para type="description">The Scopes set for the Azure AD Application.</para>
        /// <para type="description">If unset this would default to https://outlook.office.com/.default,https://outlook.office.com/SMTP.Send</para> 
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Scopes set for the Azure AD Application")]
        public String[] Scopes = {
            //"https://graph.microsoft.com/Mail.ReadWrite",
            //"https://graph.microsoft.com/Mail.Send",
            //"https://ps.outlook.com/Exchange.ManageAsApp",
            //"https://ps.outlook.com/full_access_as_app",
            //"https://ps.outlook.com/IMAP.AccessAsApp",
            //"https://ps.outlook.com/POP.AccessAsApp",
            "https://outlook.office.com/.default"
        };

        /// <summary>
        /// <para type="description">The Certificate to use for authentication.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Certificate to use for authentication.")]
        public X509Certificate2 Certificate = null;

        /// <summary>
        /// <para type="description">The Path were the PFX file is stored.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Path were the PFX file is stored.")]
        public String CertificatePath = String.Empty;

        /// <summary>
        /// <para type="description">The Password of the PFX file.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Password of the PFX file.")]
        public String CertificatePassword = String.Empty;

        protected override void BeginProcessing()
        {
            WriteVerbose(String.Format("Entering cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }

        protected override void ProcessRecord()
        {
            if (!String.IsNullOrEmpty(ClientId) && Guid.TryParse(ClientId, out Guid _))
            {
                WriteVerbose(String.Format("ClientId Provided, this is <{0}>", ClientId));
            }
            else
            {
                throw new Exception(String.Format("The ClientId <{0}> is not a GUID", ClientId));
            }

            if (!String.IsNullOrEmpty(TenantId) && Guid.TryParse(TenantId, out Guid _))
            {
                WriteVerbose(String.Format("TenantId Provided, this is <{0}>", TenantId));
            }
            else
            {
                throw new Exception(String.Format("The TenantId <{0}> is not a GUID", TenantId));
            }

            if (Scopes != null && Scopes.Length > 0)
            {
                foreach (String Scope in Scopes)
                {
                    WriteVerbose(String.Format("Scope provided: <{0}>", Scope));
                }
            }

            if ((Certificate == null) && (String.IsNullOrEmpty(CertificatePath) || String.IsNullOrEmpty(CertificatePassword)))
            {
                throw new Exception("Either Certificate or CertificatePath/CertificatePassword must be provided");
            }
            if (!String.IsNullOrEmpty(CertificatePath) && String.IsNullOrEmpty(CertificatePassword))
            {
                throw new Exception("CertificatePassword must be provided when CertificatePath is set");
            }
            if (!String.IsNullOrEmpty(CertificatePath) && !FileUtilities.IsValidPath(CertificatePath))
            {
                throw new Exception(String.Format("The path <{0}> appears to be invalid", CertificatePath));
            }
            if (!String.IsNullOrEmpty(CertificatePath) && !File.Exists(CertificatePath))
            {
                throw new Exception(String.Format("The file <{0}> cannot be found", CertificatePath));
            }
            if (!String.IsNullOrEmpty(CertificatePath) && Path.GetExtension(CertificatePath) != ".pfx")
            {
                throw new Exception(String.Format("The file <{0}> must be a password-protected PFX", CertificatePath));
            }

            AuthenticationResult authenticationResult;
            DateTime now = DateTime.UtcNow;
            X509Certificate2 authenticationCertificate = null;

            try
            {
                if (Certificate != null)
                {
                    authenticationCertificate = Certificate;
                    WriteVerbose(String.Format("Certificate parsed successfully from input Certificate."));
                }
                else
                {
                    FileStream certFile = File.OpenRead(CertificatePath);
                    byte[] certBytes = new byte[certFile.Length];
                    certFile.Read(certBytes, 0, (int)certFile.Length);
                    authenticationCertificate = new X509Certificate2(certBytes, CertificatePassword, X509KeyStorageFlags.Exportable | X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);
                    WriteVerbose(String.Format("Certificate loaded successfully from CertificatePath using CertificatePassword."));
                }
                WriteVerbose(String.Format("Thumbprint is: <{0}>. Subject is: <{1}>", authenticationCertificate.Thumbprint, authenticationCertificate.Subject));
            }
            catch (Exception ex)
            {
                WriteVerbose(String.Format("Data................: {0}", ex.Data));
                WriteVerbose(String.Format("HelpLink............: {0}", ex.HelpLink));
                WriteVerbose(String.Format("HResult.............: {0}", ex.HResult));
                WriteVerbose(String.Format("InnerException......: {0}", ex.InnerException));
                WriteVerbose(String.Format("Message.............: {0}", ex.Message));
                WriteVerbose(String.Format("Source..............: {0}", ex.Source));
                throw new Exception("Unable to Read Certificate File with the given CertificatePasword");
            }

            try
            {
                ConfidentialClientApplication confidentialApp = (ConfidentialClientApplication)ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                    .WithCertificate(authenticationCertificate)
                    .Build();

                authenticationResult = confidentialApp.AcquireTokenForClient(Scopes).ExecuteAsync().Result;
                WriteVerbose(String.Format("Application Access Token issued and expires expires at <{0}> UTC", authenticationResult.ExpiresOn.ToString("yyyy-MM-MMTHH:mm:ss")));
            }
            catch (Exception ex)
            {
                WriteVerbose(String.Format("Data................: {0}", ex.Data));
                WriteVerbose(String.Format("HelpLink............: {0}", ex.HelpLink));
                WriteVerbose(String.Format("HResult.............: {0}", ex.HResult));
                WriteVerbose(String.Format("InnerException......: {0}", ex.InnerException));
                WriteVerbose(String.Format("Message.............: {0}", ex.Message));
                WriteVerbose(String.Format("Source..............: {0}", ex.Source));
                throw ex;
            }

            JwtSecurityTokenHandler jwtHandler = new JwtSecurityTokenHandler();

            JwtSecurityToken jwtToken = jwtHandler.ReadJwtToken(authenticationResult.AccessToken);
            StringBuilder stringBuilder = new StringBuilder(1024);

            stringBuilder.AppendLine("====================== JSON WEB TOKEN DECODE =======================");
            stringBuilder.AppendLine("Header:");
            stringBuilder.AppendLine("{");
            foreach (var header in jwtToken.Header)
            {
                stringBuilder.AppendLine("\t" + '"' + header.Key + "\":\"" + header.Value + "\",");
            }
            stringBuilder.AppendLine("}");

            stringBuilder.AppendLine("Payload:");
            stringBuilder.AppendLine("{");
            foreach (var claim in jwtToken.Claims)
            {
                if (claim.Type == "iat" || claim.Type == "nbf" || claim.Type == "exp")
                {
                    DateTimeOffset dateTimeOffset = DateTimeOffset.FromUnixTimeSeconds(long.Parse(claim.Value));
                    DateTime dateTime = dateTimeOffset.UtcDateTime;
                    stringBuilder.AppendLine("\t" + '"' + claim.Type + "\":\"" + dateTime.ToString("yyyy-MM-MMTHH:mm:ss") + " UTC\",");
                }
                else
                {
                    stringBuilder.AppendLine("\t" + '"' + claim.Type + "\":\"" + claim.Value + "\",");
                }
            }
            stringBuilder.AppendLine("}");
            stringBuilder.AppendLine("====================================================================");

            WriteVerbose(stringBuilder.ToString());

            WriteObject(authenticationResult.AccessToken);
            return;
        }

        protected override void EndProcessing()
        {
            WriteVerbose(String.Format("Exiting cmdlet on <{0}>", DateTime.Now.ToString("yyyy-MM-MMTHH:mm:ss")));
        }
    }
}
