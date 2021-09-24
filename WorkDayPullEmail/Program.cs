using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkDayPullEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            string dropAttachments = @"\\cmmc\depart\IM\SoftwareDevelopers\WorkDay\DropAttachments\";

            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2016);
                service.Credentials = new WebCredentials("workdaytestemailpull@luriechildrens.org", "Cmmx0827");
                //service.AutodiscoverUrl("libarra@luriechildrens.org");
                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                //service.Url = new Uri("https://webmail.luriechildrens.org/EWS/Exchange.asxm");

                ItemView view = new ItemView(10);
                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(10));

                //service.LoadPropertiesForItems(findResults, PropertySet.FirstClassProperties); // Didn't work so remove if not using

                if (findResults != null && findResults.Items != null && findResults.Items.Count > 0)
                    foreach (Item item in findResults.Items)
                    {

                        EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments, EmailMessageSchema.From));

                        //ExtractEntities(service, message.Id); //didn't get me From address but does get URLs
                        if (message.HasAttachments)
                        {
                            foreach (Attachment attachment in message.Attachments)
                            {

                                //check if has attachments

                                if (attachment is FileAttachment)
                                {
                                    FileAttachment fileAttachment = attachment as FileAttachment;


                                    if (fileAttachment.Name.Contains(".doc") || fileAttachment.Name.Contains(".xls"))
                                    {
                                        //Move to Review folder
                                        MoveToFolder(service, message, "Review");

                                    }
                                    else
                                    {
                                        //string timeStamp = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Second.ToString() + "_" + DateTime.Now.Millisecond.ToString();
                                        string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssms");

                                        string fileExt = Path.GetExtension(fileAttachment.Name);
                                        string fielNameNoExt = fileAttachment.Name.Replace(fileExt, "");
                                        
                                        //replacing @ symbol that we're pulling from email address
                                        //Move to network path for pick up for SFTP to WorkDay
                                        fileAttachment.Load(dropAttachments + message.From.Address.Replace("@","-") + "-" + MakeValidFileName(fielNameNoExt) + "-" + timeStamp + fileExt); //saves file to location specified.  
                                       
                                    }

                                    Console.WriteLine("Attachment name: " + fileAttachment.Name);
                                }
                                else
                                {

                                    Console.Write("What do we do here?");
                                }
                            }
                            //Move after processing all attachments .
                            MoveToFolder(service, message, "Processed");
                        }
                        else
                        {
                            //no attachments found move to Review folder
                            MoveToFolder(service, message, "Review");
                        }

                        Console.WriteLine(item.Subject);

                    }
                else
                    Console.WriteLine("no items");
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
            //Console.ReadLine();
        }

        private static string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "_");
        }

        public static void MoveToFolder(ExchangeService service, EmailMessage mail, string folderName)
        {
            Folder rootfolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
            rootfolder.Load();
            Folder foundFolder = rootfolder.FindFolders(new FolderView(100)).FirstOrDefault(x => x.DisplayName == folderName);
            if (foundFolder == default(Folder))
            {
                throw new DirectoryNotFoundException(string.Format("Could not find folder {0}.", folderName));
            }

            mail.Move(foundFolder.Id);
        }

        public static void ExtractEntities(ExchangeService service, ItemId ItemId)
        {
            // Create a property set that limits the properties returned 
            // by the Bind method to only those that are required.
            PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.EntityExtractionResult);
            // Get the item from the server.
            // This method call results in an GetItem call to EWS.
            Item item = Item.Bind(service, ItemId, propSet);
            Console.WriteLine("The following entities have been extracted from the message:");
            Console.WriteLine(" ");
            // If address entities are extracted from the message, print the results.
            if (item.EntityExtractionResult != null)
            {
                if (item.EntityExtractionResult.Addresses != null)
                {
                    Console.WriteLine("--------------------Addresses---------------------------");
                    foreach (AddressEntity address in item.EntityExtractionResult.Addresses)
                    {
                        Console.WriteLine("Address: {0}", address.Address);
                    }
                    Console.WriteLine(" ");
                }
                // If contact entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.Contacts != null)
                {
                    Console.WriteLine("--------------------Contacts----------------------------");
                    foreach (ContactEntity contact in item.EntityExtractionResult.Contacts)
                    {
                        Console.WriteLine("Addresses:       {0}", contact.Addresses);
                        Console.WriteLine("Business name:   {0}", contact.BusinessName);
                        Console.WriteLine("Contact string:  {0}", contact.ContactString);
                        Console.WriteLine("Email addresses: {0}", contact.EmailAddresses);
                        Console.WriteLine("Person name:     {0}", contact.PersonName);
                        Console.WriteLine("Phone numbers:   {0}", contact.PhoneNumbers);
                        Console.WriteLine("URLs:            {0}", contact.Urls);
                    }
                    Console.WriteLine(" ");
                }
                // If email address entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.EmailAddresses != null)
                {
                    Console.WriteLine("--------------------Email addresses---------------------");
                    foreach (EmailAddressEntity email in item.EntityExtractionResult.EmailAddresses)
                    {
                        Console.WriteLine("Email addresses: {0}", email.EmailAddress);
                    }
                    Console.WriteLine(" ");
                }
                // If meeting suggestion entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.MeetingSuggestions != null)
                {
                    Console.WriteLine("--------------------Meeting suggestions-----------------");
                    foreach (MeetingSuggestion meetingSuggestion in item.EntityExtractionResult.MeetingSuggestions)
                    {
                        Console.WriteLine("Meeting subject:  {0}", meetingSuggestion.Subject);
                        Console.WriteLine("Meeting string:   {0}", meetingSuggestion.MeetingString);
                        foreach (EmailUserEntity attendee in meetingSuggestion.Attendees)
                        {
                            Console.WriteLine("Attendee name:    {0}", attendee.Name);
                            Console.WriteLine("Attendee user ID: {0}", attendee.UserId);
                        }
                        Console.WriteLine("Start time:       {0}", meetingSuggestion.StartTime);
                        Console.WriteLine("End time:         {0}", meetingSuggestion.EndTime);
                        Console.WriteLine("Location:         {0}", meetingSuggestion.Location);
                    }
                    Console.WriteLine(" ");
                }
                // If phone number entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.PhoneNumbers != null)
                {
                    Console.WriteLine("--------------------Phone numbers-----------------------");
                    foreach (PhoneEntity phone in item.EntityExtractionResult.PhoneNumbers)
                    {
                        Console.WriteLine("Original phone string:  {0}", phone.OriginalPhoneString);
                        Console.WriteLine("Phone string:           {0}", phone.PhoneString);
                        Console.WriteLine("Type:                   {0}", phone.Type);
                    }
                    Console.WriteLine(" ");
                }
                // If task suggestion entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.TaskSuggestions != null)
                {
                    Console.WriteLine("--------------------Task suggestions--------------------");
                    foreach (TaskSuggestion task in item.EntityExtractionResult.TaskSuggestions)
                    {
                        foreach (EmailUserEntity assignee in task.Assignees)
                        {
                            Console.WriteLine("Assignee name:    {0}", assignee.Name);
                            Console.WriteLine("Assignee user ID: {0}", assignee.UserId);
                        }
                        Console.WriteLine("Task string:      {0}", task.TaskString);
                    }
                    Console.WriteLine(" ");
                }
                // If URL entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.Urls != null)
                {
                    Console.WriteLine("--------------------URLs--------------------------------");
                    foreach (UrlEntity url in item.EntityExtractionResult.Urls)
                    {
                        Console.WriteLine("URL: {0}", url.Url);
                    }
                    Console.WriteLine(" ");
                }
            }
            // If no entities are extracted from the message, print the result.
            else if (item.EntityExtractionResult == null)
            {
                Console.WriteLine("No entities extracted");
            }
        }
    }
}
