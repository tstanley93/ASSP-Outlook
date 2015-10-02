using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace ASSP_Outlook
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton buttonOne;
        public Outlook.Explorer currentExplorer = null;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            //AddMenuBar();
            currentExplorer = this.Application.ActiveExplorer();
            Outlook.Selection selecteditems = currentExplorer.Selection;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        private void AddMenuBar()
        {
            try
            {
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                if (newMenuBar != null)
                {
                    newMenuBar.Caption = "New Menu";
                    buttonOne = (Office.CommandBarButton)newMenuBar.Controls.
                    Add(Office.MsoControlType.msoControlButton, missing,
                        missing, 1, true);
                    buttonOne.Style = Office.MsoButtonStyle.
                        msoButtonIconAndCaption;
                    buttonOne.Caption = "Button One";
                    buttonOne.FaceId = 65;
                    buttonOne.Tag = "c123";
                    //buttonOne.Click += new Office._CommandBarButtonEvents_ClickEventHandler(buttonOne_Click);
                    newMenuBar.Visible = true;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        //private void buttonOne_Click(Office.CommandBarButton ctrl, ref bool cancel)
        //{
        //    System.Windows.Forms.MessageBox.Show("You clicked: " + ctrl.Caption,
        //        "Custom Menu", System.Windows.Forms.MessageBoxButtons.OK);
        //}



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
