/*
 * Email.cs
 *
 * Craig Nobili
 */

using System;
using System.Net;
using System.Net.Mail;

public class Email
{
  public const String mailServer = "mr1.hsys.local";

  public static void Main(String[] args)
  {
     String[] addresses = new String[] {"craig.nobili@johnmuirhealth.com"};

     SendMail("ClarityMonitor@johnmuirhealth.com", addresses, "test", "<a href='http://www.abc.com' target='_blank'>www.abc.com</a>");

  } // Main()

  public static void SendMail(String from, String[] to, String subject, String body)
  {
    MailMessage msg = new MailMessage();
    SmtpClient SmtpServer = new SmtpClient(mailServer);

    msg.From = new MailAddress(from);
    foreach (String s in to)
    {
      msg.To.Add(s);
		}

    msg.Subject = subject;
    msg.IsBodyHtml = true;
    msg.Body = body;

    SmtpServer.Port = 25;
    SmtpServer.Send(msg);

  } // SendMail()

} // class Email