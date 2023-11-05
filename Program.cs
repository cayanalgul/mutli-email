﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Remoting.Contexts;

namespace MultiEmail
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = "filPath";
            //We open the excel file and catch the rows and columns used from the 1st sheet
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range cell = worksheet.UsedRange;

            //Foreach loop to catch values
            foreach (var item in cell.Value2)
            {
                if (item != null)
                {
                    try
                    {
                        SendMail(item);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Check the excel file it could be FormatException");
                    }
                }
                else
                {
                    Console.WriteLine("Cell is null!!!");
                }
            }
        }

        public static void SendMail(string mail)
        {
            //Send mail here

            string to = mail;
            string from = "yourMail";

            MailMessage message = new MailMessage(from, to);

            message.Attachments.Add(new Attachment(@"attachFilePath"));
            message.Subject = "subject";
            message.Body = "content";

            SmtpClient client = new SmtpClient();

            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("yourMail", "googleAppPassword");
            client.EnableSsl = true;

            try
            {
                client.Send(message);
                Console.WriteLine("Sent to this email {0}", mail);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error" + " " + ex.ToString());

            }
        }
    }
}
