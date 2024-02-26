using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace TimeStamper
{
    public partial class ThisAddIn
    {
        private Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            MessageBox.Show("Outlook Time Stamp tool has initialized!");
            Outlook.MAPIFolder calFolder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            items = calFolder.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
        }

        void Items_ItemAdd(object Item)
        {
            if (Item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)Item;
                //keep referring to documentation for item-adding
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
