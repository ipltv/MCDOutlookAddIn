using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace MCDOutlookAddIn
{
    public partial class MCDRibbon
    {
        private void MCDRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void StoreNumber_Click(object sender, RibbonControlEventArgs e)
        {
            Inspector ins = Globals.ThisAddIn.Application.ActiveInspector();
            try
            {
                string selectedText = ins.WordEditor.Application.Selection.Text;
                string storeNameByRequest = Globals.ThisAddIn.storeNameRequster.GetStoreNameByNumber(selectedText.Trim());
                if (storeNameByRequest != string.Empty)
                {
                    ins.WordEditor.Application.Selection.Text = selectedText.Trim() + " " + storeNameByRequest + selectedText.Substring(selectedText.Length - 1);
                }
            }
            catch (System.Exception er)
            {

                MessageBox.Show(er.ToString());
            }
        }

    }
}
