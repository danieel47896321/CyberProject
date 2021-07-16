using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CyberProject
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public int Stars { get; set; }
        public static List<OutlookEmails> ReadMailItems()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                foreach (MailItem item in mailItems)
                {
                    emailDetails = new OutlookEmails();
                    emailDetails.EmailFrom = item.SenderEmailAddress;
                    emailDetails.EmailSubject = item.Subject;
                    emailDetails.EmailBody = item.Body;
                    emailDetails.Stars = -1;
                    if (emailDetails.EmailBody.Contains("www.")|| emailDetails.EmailBody.Contains("https://") || emailDetails.EmailBody.Contains("http://"))
                    {
                        string htmlCode;
                        var links = emailDetails.EmailBody.Split("\t\n ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(s => s.StartsWith("https://") || s.StartsWith("www.") || s.StartsWith("http://"));
                        foreach (string s in links)
                            using (WebClient client = new WebClient())
                            {
                                string temp;
                                int count = 0;
                                if (s.Contains("https://"))
                                    count += 8;
                                if (s.Contains("http://"))
                                    count += 7;
                                if (s.Contains("www."))
                                    count += 4;
                                temp = s.Substring(count, s.Length-count);
                                try
                                {
                                    htmlCode = client.DownloadString("https://www.urlvoid.com/scan/" + temp + "/");
                                    MatchCollection ms = Regex.Matches(htmlCode, @">[0-9]+/34<");
                                    string testMatch = ms[0].Value.ToString();
                                    emailDetails.Stars = Int32.Parse(testMatch.Substring(1, testMatch.Length - 5));
                                }
                                catch
                                {
                                    emailDetails.Stars = -2;
                                }
                                
                            }
                    }
                    listEmailDetails.Add(emailDetails);
                    RelaseComObject(item);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                RelaseComObject(mailItems);
                RelaseComObject(inboxFolder);
                RelaseComObject(outlookNamespace);
                RelaseComObject(outlookApplication);
            }
            return listEmailDetails;
        }
        private static void RelaseComObject(object obj)
        {
            if (obj != null) {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}