using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System;


namespace SMTP_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MailBox();
        }

        private void MailBox()
        {
            Microsoft.Office.Interop.Outlook.Application ol = new Microsoft.Office.Interop.Outlook.Application();
           var AddressEntries = ol.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
            textBox1.Text =  AddressEntries.PrimarySmtpAddress;
            textBox2.Text = AddressEntries.PrimarySmtpAddress;

        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            //makes new mail message
            MailMessage mail = new MailMessage();
            //Sender and receiver
            mail.From = new MailAddress(textBox1.Text);
            mail.To.Add(new MailAddress(textBox2.Text));
            //Subject and body of message
            mail.Subject = "This is just a Test";
            mail.Body = "Welcome to the mail Test isnt this fun";
          
            //Try Catch to send message
            try
            {
                //Setting SMTP server and Client
                SmtpClient smtp = new SmtpClient(Properties.Settings.Default.MailSMTP,25);
                smtp.Send(mail);//Sends emal
            }
            catch (System.Exception a )
            {
                //Shows Message box to show exception
                MessageBox.Show("Copy and Paste this in an email to the Helpdesk: \n\n" + a.ToString());
               
            }
        }
    }
}
