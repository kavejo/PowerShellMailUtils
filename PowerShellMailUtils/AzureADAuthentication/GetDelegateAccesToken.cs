using Microsoft.Identity.Client;
using SyntheticTransactionsForExchange.Utilities;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace SyntheticTransactionsForExchange.AzureADAuthentication
{
    [Cmdlet("Get", "DelegateAccesToken")]
    [OutputType(typeof(String))]
    [CmdletBinding()]

    public class GetDelegateAccesToken : Cmdlet
    {
        /// <summary>
        /// <para type="description">The ClientId set on the Azure AD Application.</para>
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = @"The ClientId set on the Azure AD Application")]
        public String ClientId = String.Empty;

        /// <summary>
        /// <para type="description">The RedirectUri set on the Azure AD Application. Mandatory if the scopes include GRAPH resources.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The RedirectUri set on the Azure AD Application. Mandatory if the scopes include GRAPH resources")]
        public String RedirectUri = String.Empty;

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
            //"https://graph.microsoft.com/EWS.AccessAsUser.All",
            //"https://graph.microsoft.com/IMAP.AccessAsUser.All",
            //"https://graph.microsoft.com/Mail.ReadWrite.Shared",
            //"https://graph.microsoft.com/Mail.Send",
            //"https://graph.microsoft.com/POP.AccessAsUser.All",
            //"https://graph.microsoft.com/SMTP.Send",
            "https://outlook.office.com/EWS.AccessAsUser.All",
            "https://outlook.office.com/SMTP.Send"
        };

        /// <summary>
        /// <para type="description">The Credentials of the user for which the token is to be requested.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The Credentials of the user for which the token is to be requested.")]
        public PSCredential Credentials = null;

        /// <summary>
        /// <para type="description">The UnserName of the user for which the token is to be requested.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"The UnserName of the user for which the token is to be requested")]
        public String UnserName = String.Empty;

        /// <summary>
        /// <para type="description">Password for the account for which to get the token.</para>
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = @"Password for the account for which to get the token.")]
        public String Password = String.Empty;

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

            if (Scopes.FirstOrDefault(s => s.Contains("graph")) != null && String.IsNullOrEmpty(RedirectUri))
            {
                throw new Exception(String.Format("RedirectUri is mandatory if Graph is to be targeted. Uri set to <{0}>", RedirectUri));
            }

            if ((Credentials == null) && (String.IsNullOrEmpty(UnserName) || String.IsNullOrEmpty(Password)))
            {
                throw new Exception("Either Credentials or UnserName/Password must be provided");
            }

            if (!String.IsNullOrEmpty(UnserName) && !RegexUtilities.IsValidEmail(UnserName))
            {
                throw new Exception(String.Format("The UnserName <{0}> is not a valid UPN in the SMTP form", UnserName));
            }

            if (!String.IsNullOrEmpty(UnserName) && String.IsNullOrEmpty(Password))
            {
                throw new Exception(String.Format("Password cannot be empty if UnserName is set"));
            }

            if (Credentials != null)
            {
                UnserName = Credentials.UserName;
                Password = Credentials.GetNetworkCredential().Password;
                WriteVerbose(String.Format("Username and Password set from Credentials"));
            }

            //SecureString securePassword = new SecureString();
            AuthenticationResult authenticationResult;
            DateTime now = DateTime.UtcNow;

            try
            {
                //foreach (char ch in Password)
                //    securePassword.AppendChar(ch);

                PublicClientApplication publicApp = null;

                if (!String.IsNullOrEmpty(RedirectUri))
                {

                    WriteVerbose(String.Format("RedirectUri provided: <{0}>", RedirectUri));
                    publicApp = (PublicClientApplication)PublicClientApplicationBuilder
                        .Create(ClientId)
                        .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                        .WithRedirectUri(RedirectUri)
                        .Build();
                }
                else
                {
                    WriteVerbose(String.Format("Connecting without RedirectUri"));
                    publicApp = (PublicClientApplication)PublicClientApplicationBuilder
                        .Create(ClientId)
                        .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                        .Build();
                }

                authenticationResult = publicApp.AcquireTokenByUsernamePassword(Scopes, UnserName, Password).ExecuteAsync().Result;
                WriteVerbose(String.Format("Delegate Access Token issued for <{0}> and expires at <{1}> UTC", authenticationResult.Account.Username, authenticationResult.ExpiresOn.ToString("yyyy-MM-MMTHH:mm:ss")));
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
