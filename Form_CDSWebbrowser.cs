using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCHelper_v001
{
    public partial class Form_CDSWebbrowser : Form
    {

        private string userName = "stdntuvmath@gmail.com";
        private string password = "yrthsa12";
        private string cdsURL = "https://www.facebook.com/";


        public Form_CDSWebbrowser()
        {
            InitializeComponent();
        }

        private void InvisibleWebBrowser_Load(object sender, EventArgs e)
        {
            webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(checkDocument);
            webBrowser1.Navigate(cdsURL);

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            webBrowser1 = sender as WebBrowser;


        }
        private void checkDocument(object sender, WebBrowserDocumentCompletedEventArgs e)
        {


            try
            {
                HtmlDocument doc = webBrowser1.Document;
                doc.GetElementById("email").InnerText = userName;
                doc.GetElementById("pass").InnerText = password;
            }
            catch (System.ArgumentException ex)
            {
                MessageBox.Show("Method: HtmlDocument.GetElementById()\r\rCan't get element by ID.\r\r"+ex,"",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
           

           // doc.GetElementById("u_0_2").InvokeMember("click");

           // webBrowser1.Document.GetElementById("Ticket_Seach_Filters").InvokeMember("click");
            //MessageBox.Show("Posted");

        }



        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
           
        }
    }
}
