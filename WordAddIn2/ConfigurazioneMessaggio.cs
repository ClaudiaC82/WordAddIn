using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MailerUtilities;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace WordAddIn2
{
    public partial class ConfigurazioneMessaggio : Form
    {
        string messaggio = "";
        public ConfigurazioneMessaggio()
        {
            InitializeComponent();
        }

        private void ConfigurazioneMessaggio_Load(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document documento = Globals.ThisAddIn.Application.ActiveDocument;
            messaggio = documento.Content.Text;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            XLSX excel = new XLSX();
            string base_directory = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = ConfigReader.PersistableObject.Load<ConfigReader.Configuration>(String.Format("{0}\\config.xml", base_directory)).excel_path;
            excel.elaboraExcel(filePath);//"D:\\provaExcel.xlsx"
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Microsoft.Office.Interop.Outlook.
            Outlook.Application olkApp1 = new Outlook.Application();
            Outlook.MailItem olkMail1 =  (MailItem)olkApp1.CreateItem(OlItemType.olMailItem);
            olkMail1.To = "claudia_cassioli@libero.it";
            olkMail1.CC = "";
            olkMail1.Subject = "test subject";
            olkMail1.Body = "test test test";
            ////olkMail1.Attachments.Add(AssignNoteFilePath,
            ////    Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1,
            ////        "Assignment_note");
            olkMail1.Send();

        }
    }
}
