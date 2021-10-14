using System;
using System.Threading.Tasks;

namespace GraphEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            GraphEmailService emailInstance = new GraphEmailService();

            //emailInstance.SendEmailAsync().GetAwaiter().GetResult();

            //emailInstance.GetInboxMessgesAsync().GetAwaiter().GetResult();

            //emailInstance.SearchEmailBySubjectAsync().GetAwaiter().GetResult();

            //emailInstance.GetMailFoldersAsync().GetAwaiter().GetResult();

            //emailInstance.GetMessageIdBasedOnSubjectContentAsync("kkk").GetAwaiter().GetResult();

            emailInstance.ReplyToEmailAsync().GetAwaiter().GetResult();

            Console.WriteLine("Completed");
            Console.ReadKey();
        }
    }
}
