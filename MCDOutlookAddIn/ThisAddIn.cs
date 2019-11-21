using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.IO;

namespace MCDOutlookAddIn
{
    public partial class ThisAddIn
    {
        public StoreNameClaim storeNameRequster;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DirectoryInfo AppDir = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MCDOutlookAddIn"));
            if (!AppDir.Exists)
            {
                MessageBox.Show("Application directory not found and will be created");
                AppDir.Create();
            }
            try
            {
                storeNameRequster = new StoreNameClaim(Path.Combine(AppDir.FullName, @"storeslist.txt"));
            }
            catch (Exception ex)
            {
                
                MessageBox.Show("Current File: " + Path.Combine(AppDir.FullName, @"storeslist.txt") + "\n" + ex.ToString());
            }

            Outlook.Items items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;
            items.ItemAdd += Items_ItemAdd;
        }

        private void Items_ItemAdd(object Item)
        {

            try
            {
                Outlook.MailItem mailItem = (Outlook.MailItem)Item;
                if (mailItem.Subject.Contains("NewPOS license file"))
                {
                    DirectoryInfo storeDir = GetStoreDirectory(mailItem.Body);
                    Outlook.Attachment licenseFile = mailItem.Attachments[0];
                    string filePath = Path.Combine(storeDir.FullName, licenseFile.FileName);
                    if (File.Exists(filePath)) File.Delete(filePath);
                    licenseFile.SaveAsFile(filePath);
                }
            }
            catch
            {
                MessageBox.Show($"Не удалось сохранить файл лицензии из письма {((Outlook.MailItem)Item).Subject}");
                return;
            }
           
        }
        private DirectoryInfo GetStoreDirectory(string body)
        {
            string[] strings = body.Split(Environment.NewLine.ToCharArray());
            string storeNumber = string.Empty, storeName = string.Empty;

            foreach(string item in strings)
            {
                if(item.Contains("Store Name"))
                {
                    storeName = item.Split(':')[1];  
                }
                if(item.Contains("World ID"))
                {
                    storeNumber = item.Split(':')[1];
                }
            }
            if (string.IsNullOrEmpty(storeNumber) || string.IsNullOrEmpty(storeName)) throw new InvalidOperationException("Не удалось определить имя или номер ПБО из лицензионного сообщения");
            DirectoryInfo result = new DirectoryInfo(Path.Combine(@"G:\POOLS\IS\STORE SYSTEM\Licenses\POS v6",storeNumber + " " + storeName));
            if (!result.Exists) result.Create();
            return result;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Примечание. Outlook больше не выдает это событие. Если имеется код, который 
            //    должно выполняться при завершении работы Outlook, см. статью на странице https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private void RDILicenseMove(object item)
        {

        }
    }
}
