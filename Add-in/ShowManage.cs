using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Add_in
{
    [ComVisible(true)]

    [ClassInterface(ClassInterfaceType.AutoDual)]

    internal static class ShowManage
    {
        static CustomTaskPane ctpConfig;
       
        public static void ShowCTPSetting()
        {
            if (ctpConfig == null)
            {
                try
                {
                    // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                    ctpConfig = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(frmConfig), "Cấu hình - AccNet UX");
                    ctpConfig.Width = 320;
                    ctpConfig.Visible = true;
                    ctpConfig.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    ctpConfig.DockPositionStateChange += ctp_DockPositionStateChange;
                    ctpConfig.VisibleStateChange += ctp_VisibleStateChange;
                }
                catch (Exception ex)
                {
                    throw;
                }
              
            }
            else
            {
                // Just show it again
                ctpConfig.Visible = true;
            }
        }
        static void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            //MessageBox.Show("Visibility changed to " + CustomTaskPaneInst.Visible);
        }

        static void ctp_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            //((MyUserControl)ctp.ContentControl).TheLabel.Text = "Moved to " + CustomTaskPaneInst.DockPosition.ToString();
        }

        public static void DeleteCTPSetting()
        {
            if (ctpConfig != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctpConfig.Delete();
                ctpConfig = null;
            }
        }

    }
}
