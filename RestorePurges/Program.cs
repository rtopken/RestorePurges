using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using static RestorePurges.PropSet;


namespace RestorePurges
{
    class Program
    {
        public static ExchangeService exService;
        public static string strClientID = "75b8688c-18c7-4e5a-8831-b5dc18dd125a";
        public static string strRedirURI = "https://RestorePurges";
        public static string strAuthCommon = "https://login.microsoftonline.com/common";
        public static string strSrvURI = "https://outlook.office365.com";
        public static string strDisplayName = "";
        public static string strSMTPAddr = "";
        public static int cItems = 0;
        public static Folder fldCal = null;
        public static Folder fldPurges = null;
        public static bool bStartDate = false;
        public static DateTime dtStartDate = DateTime.MinValue;


        static void Main(string[] args)
        {
            string strAcct = "";
            string strTenant = "";
            string strEmailAddr = "";
            bool bMailbox = false;
            NameResolutionCollection ncCol = null;
            List<string> strCalList = new List<string>();
            List<Item> CalItems = null;
            List<Item> PurgeItems = null;
            int cRestoredItems = 0;
            string strStartDate = "";

            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].ToUpper() == "-M" || args[i].ToUpper() == "/M") // mailbox mode - use impersonation to get to another mailbox
                    {
                        if (args[i + 1].Length > 0)
                        {
                            strEmailAddr = args[i + 1];
                            bMailbox = true;
                        }
                        else
                        {
                            Console.WriteLine("Please enter a valid SMTP address for the mailbox.");
                            ShowHelp();
                            return;
                        }
                    }
                    if (args[i].ToUpper() == "-S" || args[i].ToUpper() == "/S") // Choose a Start Date instead of using 24 hours
                    {
                        if (args[i + 1].Length > 0)
                        {
                            strStartDate = args[i + 1];
                            bStartDate = true;
                        }
                        else
                        {
                            Console.WriteLine("Please enter a valid Date.");
                            ShowHelp();
                            return;
                        }
                    }
                    if (args[i].ToUpper() == "-?" || args[i].ToUpper() == "/?") // display command switch help
                    {
                        ShowInfo();
                        ShowHelp();
                        return;
                    }
                }
            }

            if (bStartDate)
            {
                try
                {
                    dtStartDate = DateTime.Parse(strStartDate);
                }
                catch (Exception ex)
                {
                    if (ex.Message == "String was not recognized as a valid DateTime.")
                    {
                        Console.WriteLine("Date Format was incorrect. Please enter a valid Date.\r\n");
                        ShowHelp();
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Error with Entered Date. Please try again.");
                        Console.WriteLine(ex.Message + "\r\n");
                        ShowHelp();
                        return;
                    }
                }
            }

            strStartDate = dtStartDate.ToString();

            ShowInfo();

            exService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            exService.UseDefaultCredentials = false;

            Console.Write("Press <ENTER> to enter credentials.");
            Console.ReadLine();
            Console.WriteLine();

            AuthenticationResult authResult = GetToken();
            if (authResult != null)
            {
                exService.Credentials = new OAuthCredentials(authResult.AccessToken);
                strAcct = authResult.UserInfo.DisplayableId;
            }
            else
            {
                return;
            }
            strTenant = strAcct.Split('@')[1];
            exService.Url = new Uri(strSrvURI + "/ews/exchange.asmx");

            if (bMailbox)
            {
                ncCol = DoResolveName(strEmailAddr);
                exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strEmailAddr);
            }
            else
            {
                ncCol = DoResolveName(strAcct);
            }

            if (ncCol == null)
            {
                // Didn't get a NameResCollection, so error out.
                Console.WriteLine("");
                Console.WriteLine("Exiting the program.");
                return;
            }

            if (ncCol[0].Contact != null)
            {
                strDisplayName = ncCol[0].Contact.DisplayName;
                strEmailAddr = ncCol[0].Mailbox.Address;
                Console.WriteLine("Will attempt restore from Purges folder for " + strDisplayName);
            }
            else
            {
                Console.WriteLine("Will attempt restore from Purges folder for " + strAcct);
            }

            Console.WriteLine("==============================================================\r\n");

            // Get Calendar items
            Console.WriteLine("Connecting to Calendar and retreiving items.");
            CalItems = GetItems(exService, "Calendar");

            if (CalItems != null)
            {
                string strCount = CalItems.Count.ToString();
                Console.WriteLine("Creating Calendar items list for " + strCount + " items.\r\n");
                foreach (Appointment appt in CalItems)
                {
                    strCalList.Add(GetPropsLine(appt)); //strGlobalObjID + "," + strSubject + "," + strStartWhole + "," + strEndWhole + "," + strOrganizerAddr + "," + strRecurring
                }
            }
            else
            {
                return;
            }

            // Get Purges items
            if (bStartDate)
            {
                Console.WriteLine("Connecting to Purges folder and retreiving purged Calendar items starting from " + strStartDate + ".");
            }
            else
            {
                Console.WriteLine("Connecting to Purges folder and retreiving purged Calendar items from the last 24 hours.");
            }
            PurgeItems = GetItems(exService, "Purges");

            if (PurgeItems != null)
            {
                if (PurgeItems.Count > 0)
                {
                    Console.WriteLine("Checking " + PurgeItems.Count.ToString() + " items from Purges against the existing Calendar items list.\r\n");
                    foreach (Appointment appt in PurgeItems)
                    {
                        string strPurged = GetPropsLine(appt);
                        bool bRestore = true;

                        foreach (string strCalItem in strCalList)
                        {
                            if (strPurged == strCalItem)
                            {
                                bRestore = false;
                            }
                        }

                        if (bRestore)
                        {
                            string strStart = "";
                            string strEnd = "";
                            try
                            {
                                strStart = appt.Start.ToString();
                            }
                            catch
                            {
                                strStart = "Cannot Retrieve or Not Set.";
                            }

                            try
                            {
                                strEnd = appt.End.ToString();
                            }
                            catch
                            {
                                strEnd = "Cannot Retrieve or Not Set.";
                            }

                            Console.WriteLine("Recovering item:");
                            Console.WriteLine(appt.Subject + " | Location: " + appt.Location + " | Start Time: " + strStart + " | End Time: " + strEnd);
                            appt.Move(fldCal.Id);
                            cRestoredItems++;
                            strCalList.Add(strPurged); // this one has been moved back, so add it to the list in case there is another older one
                        }
                    }
                }
                else
                {
                    if (bStartDate)
                    {
                        Console.WriteLine("There were no Calendar items sent to the Purges folder since " + strStartDate + ".\r\n");
                    }
                    else
                    {
                        Console.WriteLine("There were no Calendar items sent to the Purges folder in the last 24 hours.\r\n");
                    }
                }
            }
            else
            {
                return;
            }

            string strItemCount = "";
            if (cRestoredItems == 1)
            {
                strItemCount = "1 item";
            }
            else
            {
                strItemCount = cRestoredItems.ToString() + " items";
            }

            Console.WriteLine("========================");
            Console.WriteLine("Complete. Restored " + strItemCount + " from Purges back to Calendar.");
            Console.WriteLine("========================");
        }


        public static void ShowInfo()
        {
            Console.WriteLine("");
            Console.WriteLine("=============");
            Console.WriteLine("RestorePurges");
            Console.WriteLine("=============");
            Console.WriteLine("Restores Calendar Items from the Purges folder.\r\n");
        }


        public static void ShowHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("RestorePurges [-M <SMTP Address>] [-S <Date M/D/Y>] [-?]");
            Console.WriteLine("");
            Console.WriteLine("-M   [Mailbox - will connect to the mailbox and perform the restore.]");
            Console.WriteLine("-S   [Date - will use the date as a start point for searching the Purges folder.]");
            Console.WriteLine("-?   [Shows this usage information.]");
            Console.WriteLine("");
        }


        // Go connect to afolder and get the items
        public static List<Item> GetItems(ExchangeService service, string strFolder)
        {
            Folder fld = null;
            int iOffset = 0;
            int iPageSize = 500;
            bool bMore = true;
            List<Item> lItems = new List<Item>();
            FindItemsResults<Item> findResults = null;
            DateTime dtNow = DateTime.Now;
            DateTime dtBack = dtNow.AddHours(-24);
            
            if (bStartDate) // default is 24 hours ago on the date, but the user can specify longer ago than that if they wish.
            {
                dtBack = dtStartDate;
            }
            
            SearchFilter.ContainsSubstring apptFilter = new SearchFilter.ContainsSubstring(ItemSchema.ItemClass, "IPM.Appointment");
            SearchFilter.IsGreaterThanOrEqualTo modifiedFilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, dtBack);
            SearchFilter.SearchFilterCollection multiFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, apptFilter, modifiedFilter);

            if (strFolder == "Calendar")
            {
                try
                {
                    // Here's where it connects to the Calendar
                    fld = Folder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
                    fldCal = fld;
                }
                catch (ServiceResponseException ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("");
                    Console.WriteLine("Could not connect to this user's mailbox or Calendar folder.");
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                    return null;
                }
            }
            else if (strFolder == "Purges")
            {
                try
                {
                    // Here's where it connects to Purges
                    fld = Folder.Bind(service, WellKnownFolderName.RecoverableItemsPurges, new PropertySet());
                    fldPurges = fld;
                }
                catch (ServiceResponseException ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("");
                    Console.WriteLine("Could not connect to this user's mailbox or Purges folder.");
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                    return null;
                }
            }
            else
            {
                Console.WriteLine("Not Purges or Calendar - no connection.");
                return null;
            }

            // if we're in then we get here
            // creating a view with props to request / collect
            ItemView cView = new ItemView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            List<ExtendedPropertyDefinition> propSet = new List<ExtendedPropertyDefinition>();
            DoProps(ref propSet);
            cView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
            foreach (PropertyDefinitionBase pdbProp in propSet)
            {
                cView.PropertySet.Add(pdbProp);
            }

            if (strFolder == "Purges")
            {
                cView.OrderBy.Add(ItemSchema.LastModifiedTime, SortDirection.Descending);
                while (bMore)
                {
                    findResults = fld.FindItems(multiFilter, cView);

                    foreach (Item item in findResults.Items)
                    {
                        lItems.Add(item);
                    }

                    bMore = findResults.MoreAvailable;
                    if (bMore)
                    {
                        cView.Offset += iPageSize;
                    }
                }
            }
            else // Calendar folder - don't need to do the search filtering here...
            {
                while (bMore)
                {
                    findResults = fld.FindItems(cView);

                    foreach (Item item in findResults.Items)
                    {
                        lItems.Add(item);
                    }

                    bMore = findResults.MoreAvailable;
                    if (bMore)
                    {
                        cView.Offset += iPageSize;
                    }
                }
            }

            return lItems;
        }


        // Go get an OAuth token to use Exchange Online 
        private static AuthenticationResult GetToken()
        {
            AuthenticationResult ar = null;
            AuthenticationContext ctx = new AuthenticationContext(strAuthCommon);

            try
            {
                ar = ctx.AcquireTokenAsync(strSrvURI, strClientID, new Uri(strRedirURI), new PlatformParameters(PromptBehavior.Always)).Result;
            }
            catch (Exception Ex)
            {
                var authEx = Ex.InnerException as AdalException;

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error occurred during authentication with the service:");
                Console.WriteLine(authEx.HResult.ToString("X"));
                Console.WriteLine(authEx.Message);
                Console.ResetColor();
            }
            return ar;
        }


        public static NameResolutionCollection DoResolveName(string strResolve)
        {
            NameResolutionCollection ncCol = null;
            try
            {
                ncCol = exService.ResolveName(strResolve, ResolveNameSearchLocation.DirectoryOnly, true);
            }
            catch (ServiceRequestException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error when attempting to resolve the name for " + strResolve + ":");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
                return null;
            }

            return ncCol;
        }
    }
}
