using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

//dotnet build 
//dotnet run -- "47385850-adf4-4a42-9d57-87528c1223fe" "Va38Q~clNwkOraFCnCHpWpmyoxSaK6uucpSzScTJ" "11bbefd4-0c6d-4f57-b665-8595ef75b628"

namespace EWSConsoleApp3
{
internal class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)  // only one Main; add public if does not work; 
        //The addition of async and Task, Task<int> return types simplifies program code when console applications need to start and await asynchronous operations in Main.
        // async in Main means the compiler generates the boilerplate code for calling asynchronous methods in Main, so dont have to write that code myself 
        // this returns object type of System.Threading.Tasks.Task       
        {
//should add code Test if input arguments were supplied.
        
            // args must alwasy be AzureApp Client - ID Secrete TenantId
            // ex: program.exe "47385850-adf4-4a42-9d57-87528c1223fe" "Va38Q~clNwkOraFCnCHpWpmyoxSaK6uucpSzScTJ" "11bbefd4-0c6d-4f57-b665-8595ef75b628"
            // args[0]
       


            // Using Microsoft.Identity.Client
            var cca = ConfidentialClientApplicationBuilder
                .Create(args[0])
                .WithClientSecret(args[1])
                .WithTenantId(args[2])
                .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            try
            {
                // Get token
                var authResult = await cca.AcquireTokenForClient(ewsScopes)
                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
                ewsClient.ImpersonatedUserId =
                    new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "LynneR@tw26j.onmicrosoft.com");

                //Include x-anchormailbox header
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", "LynneR@tw26j.onmicrosoft.com");

               // https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-work-with-folders-by-using-ews-in-exchange
                // Create a new folder view, and pass in the maximum number of folders to return.
                FolderView view = new FolderView(100);
                // Create an extended property definition for the PR_ATTR_HIDDEN property,
                // so that your results will indicate whether the folder is a hidden folder.
                ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
                // As a best practice, limit the properties returned to only those required.
                // In this case, return the folder ID, DisplayName, and the value of the isHiddenProp
                // extended property.
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.TotalCount, FolderSchema.FolderClass, FolderSchema.WellKnownFolderName, isHiddenProp);
                // Indicate a Traversal value of Deep, so that all subfolders are retrieved.
                view.Traversal = FolderTraversal.Deep;
                // Call FindFolders to retrieve the folder hierarchy, starting with the MsgFolderRoot folder.
                // This method call results in a FindFolder call to EWS.
                FindFoldersResults findFolderResults = ewsClient.FindFolders(WellKnownFolderName.Inbox, view); 
foreach (var folder in findFolderResults)
                    if (folder.FolderClass == "IPF.Note")
                    {
                        Console.WriteLine(folder.DisplayName + " " + folder.TotalCount);
                    }

                // https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-work-with-exchange-mailbox-items-by-using-ews-in-exchange
                // Bind the Inbox folder to the service object.
                Folder inbox = Folder.Bind(ewsClient, WellKnownFolderName.Inbox);
                // The search filter to get unread email.
                SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                ItemView iView = new ItemView(100);
                // Fire the query for the unread items.
                // This method call results in a FindItem call to EWS.
                FindItemsResults<Item> findResults = ewsClient.FindItems(WellKnownFolderName.Inbox, sf, iView);

                // findResults.GetEnumerator().MoveNext();

                foreach (var Email in findResults)
                {
                    Console.WriteLine("inbox email: ");
                    EmailMessage mymessage = EmailMessage.Bind(ewsClient, Email.Id);
                    //Console.WriteLine(Email.Id);
                    Console.WriteLine("To: "+Email.DisplayTo);
                    Console.WriteLine("From: " + mymessage.From);
                    Console.WriteLine("Subject: " + Email.Subject);
                    Console.WriteLine("Has Attachments: " + Email.HasAttachments);
                    
                    // Console.WriteLine(mymessage.Body);
                    //Console.WriteLine(mymessage.TextBody.ToString()); 

                    Console.WriteLine(" ");

                    // GetMessageBodyAsHtml() how to make this work?
                }

            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }

            //if (System.Diagnostics.Debugger.IsAttached)
            //{
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            //}
        
        } 
        public static void GetMessageBodyAsHtml(Item oItem, ref string MessageBodyHtml)
        {
            string sRet = string.Empty;

            PropertySet oPropSet = new PropertySet(PropertySet.FirstClassProperties);
            oItem.Load(PropertySet.FirstClassProperties);

            PropertySet oPropSetForBodyText = new PropertySet(PropertySet.FirstClassProperties);
            oPropSetForBodyText.RequestedBodyType = BodyType.HTML;
            oPropSetForBodyText.Add(ItemSchema.Body);
            oItem.Service.ClientRequestId = Guid.NewGuid().ToString();  // Set a new GUID
            Item oItemForBodyText = Item.Bind((ExchangeService)oItem.Service, oItem.Id, (PropertySet)oPropSetForBodyText);
            oItem.Service.ClientRequestId = Guid.NewGuid().ToString();  // Set a new GUID
            oItem.Load(oPropSetForBodyText);
            sRet = oItem.Body.Text;
            MessageBodyHtml = sRet;
        }

        
    }
}

