using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CyberProject
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookEmails.ReadMailItems();
            int i = 1;
            foreach (var mail in mails)
            {
                if (mail.Stars==1 || mail.Stars==0)  // 0-1
                {
                    Console.WriteLine("Mail number " + i + " from "+ mail.EmailFrom+" is Safe "+ "\t[Blacklist Status: "+ mail.Stars+"/34]");
                    Console.WriteLine("Mail Title: " + mail.EmailSubject);
                    Console.WriteLine("Mail Body " + mail.EmailBody);
                    Console.WriteLine("");
                    i = i + 1;
                }
                else if(mail.Stars >= 2 && mail.Stars < 10)  // 2-9
                {
                    Console.WriteLine("Mail number " + i + " from " + mail.EmailFrom + " Very suspicious "+ "\t[Blacklist Status: "+ mail.Stars+" / 34]");
                    Console.WriteLine("Mail Title: " + mail.EmailSubject);
                    Console.WriteLine("Mail Body " + mail.EmailBody);
                    Console.WriteLine("");
                    i = i + 1;
                }
                else if(mail.Stars > 10) //10-34
                {
                    Console.WriteLine("Mail number " + i + " is NOT Safe!!! " + "\t[Blacklist Status: " + mail.Stars + "/34]");
                    Console.WriteLine("Mail Recieved from " + mail.EmailFrom);
                    Console.WriteLine("Mail Subject " + mail.EmailSubject);
                    Console.WriteLine("Mail Body " + mail.EmailBody);
                    Console.WriteLine("");
                    i = i + 1;
                }
                else if(mail.Stars == -1)
                {
                    Console.WriteLine("Mail number " + i + " is Safe!!! (There is not url inside the mail) ");
                    Console.WriteLine("Mail Recieved from " + mail.EmailFrom);
                    Console.WriteLine("Mail Subject " + mail.EmailSubject);
                    Console.WriteLine("Mail Body " + mail.EmailBody);
                    Console.WriteLine("");
                    i = i + 1;
                }
                else
                {
                    Console.WriteLine("Mail number " + i + " is Safe!!! (The website inside not exsist)");
                    Console.WriteLine("Mail Recieved from " + mail.EmailFrom);
                    Console.WriteLine("Mail Subject " + mail.EmailSubject);
                    Console.WriteLine("Mail Body " + mail.EmailBody);
                    Console.WriteLine("");
                    i = i + 1;
                }
            }
            Console.ReadKey();
        }
    }
}