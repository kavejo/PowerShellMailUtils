# PowerShellMailUtils

This PowerShell Binary Module conveniently offers comdlets to Send, Receive and Delete mail messages from given mail servers.
Supported protocols are GRAPH, EWS, POP, IMAP, SMTP with a variety of authentication mechanisms such as Basic, Windows Integrated, Kerberos, Modern Authentication.


## Referenced Libraries
POP, IMAP and SMTP cmdlets are built on top of [MimeKit](https://github.com/jstedfast/MimeKit) and [MailKit](https://github.com/jstedfast/MailKit).

Exchange Web Services relies on the [EWS Managed API](https://github.com/OfficeDev/ews-managed-api) while GRAPH support is provided by the [GRAPh SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet). Owner Access, Delegate Access and Application Impersonation are all supported.

Modern Authentication is built on top of [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet).


## Example Use Cases

The binary module offers two set of command types. One that includes "Monitoring" in the Cmdlet name and which allows to do integrations with Monitoring Systems such as Azure RunBooks or System Center Operations Manager. The other that does not have "Monitoring" in the name offers the ability to send, receive or delete messages via any of the supported protocols.


## Authentication

Authentication can happen via *Get-ApplicationAccesToken* if you have registered an application in Azure AD Enterprise Registration to leverage App permissions. The permissions assigned must be aligned to what the token will be used for (i.e. Receiving via GRAPH, EWS, POP, IMAP or Sending via GRAPh, EWS or SMTP). Usually this type of authentication can be performed with Certificate (PFX + Password, from Certificate Store, from Variable) or by using a Client Secret. 

Similarly, the module can retrieve a Delegate Access token by using *Get-DelegateAccesToken*; This relies on Delegate Permissions which allow the user to either be prompted for authentication or to supply Username and Password (Secure String or String) to allow fetching the token in background.

For other type of authentication, it is possible to pass a Credential Object, plain Username and Password, or use the contenxt of the logged-on user.

## Monitoring Example

Without going in detail about fetching a token, the *Send-<PROTO>MonitoringMail* allows to send a predefined email message to a given recipient. The command wuld report on how long the chosen protocl has taken to process the request. Email messages are identified by a chosen GUID that is set in the subkect.

```
SubjectGuid             : f0daaf56-82b0-4819-936e-920c1005e471
TimeSent                : 22/02/2023 13:55:49
SendingProtocol         : EWS
TimeReceived            : 01/01/0001 00:00:00
ReceivingProtocol       : None
Latency                 : 0
CmdletDuration          : 645
Status                  : Success
MailflowHeaderDataTable : {}
```

The *Get-<PROTO>MonitoringMail* command allows to retrieve the message with the specific subject from the recipient mailbox. After fetching the email the cmdlet would report on the message delivery latency, as a difference between the time stamped on the message body (sending time) and the time at which the message is received, then perform analysis on the Message Headers so to identify potential delays in transit.

```
SubjectGuid             : f0daaf56-82b0-4819-936e-920c1005e471
TimeSent                : 22/02/2023 13:55:49
SendingProtocol         : EWS
TimeReceived            : 22/02/2023 13:55:53
ReceivingProtocol       : EWS
Latency                 : 3246
CmdletDuration          : 562
Status                  : Success
MailflowHeaderDataTable : {
                           The message at hop <1> has been submitted by <MN2PR00MB0608.namprd00.prod.outlook.com> to <MN2PR00MB0608.namprd00.prod.outlook.com> via <MAPI>; This was received by the receiving host on <2023-02-02T14:55:50> with a total latency of <0> seconds.
                           The message at hop <2> has been submitted by <MN2PR00MB0608.namprd00.prod.outlook.com> to <CH2PR00MB0780.namprd00.prod.outlook.com> via <SMTP>; This was received by the receiving host on <2023-02-02T14:55:50> with a total latency of <0> seconds.
                           The message at hop <3> has been submitted by <CH2PR00MB0780.namprd00.prod.outlook.com> to <BY5PR00MB0822.namprd00.prod.outlook.com> via <HTTPS>; This was received by the receiving host on <2023-02-02T14:55:53> with a total latency of <3> seconds.
                          }
```

The Header data which is displayed in text format in the example above, is also available as an object that can be further manipulated.

```
Hop  SubmittingHost                           ReceivingHost                            Type    ReceivedTime          DelayTime
---  --------------                           -------------                            ----    ------------          ---------
  1  MN2PR00MB0608.namprd00.prod.outlook.com  MN2PR00MB0608.namprd00.prod.outlook.com  mapi    2023-02-02T14:55:50   00:00:00
  2  MN2PR00MB0608.namprd00.prod.outlook.com  CH2PR00MB0780.namprd00.prod.outlook.com  SMTP    2023-02-02T14:55:50   00:00:00
  3  CH2PR00MB0780.namprd00.prod.outlook.com  BY5PR00MB0822.namprd00.prod.outlook.com  HTTPS   2023-02-02T14:55:53   00:00:03
```

## Message Exchange Example

This has yet to be implemented, however the idea is to allow sending messages in a way that From, To, Cc, Bcc, Subject and Body can be provided as input parameter.
At this time this has yet to be developed.
