using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LSExtensionWindowLib;
using LSSERVICEPROVIDERLib;
using Patholab_DAL_V1;
using Patholab_Common;
using InspectionPlan;
using System.Runtime.InteropServices;
using Spire.Doc;
using System.IO;

namespace DisplayInspection
{
    [ComVisible(true)]
    [ProgId("InspectionPlan.InspectionPlan")]
    public partial class UserControl1: UserControl, IExtensionWindow
    {

        #region Private members

        private INautilusProcessXML xmlProcessor;
        private INautilusUser _ntlsUser;
        private IExtensionWindowSite2 _ntlsSite;
        private INautilusServiceProvider sp;
        private INautilusDBConnection _ntlsCon;

        #endregion

        public UserControl1()
        {
            try
            {
                InitializeComponent();
                this.Disposed += InspectionPlanForm_Disposed;
                BackColor = Color.FromName("Control");
                this.Dock = DockStyle.Fill;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        void InspectionPlanForm_Disposed(object sender, EventArgs e)
        {
            GC.Collect();
        }

        public bool CloseQuery()
        {
            DialogResult res = MessageBox.Show(@"?האם אתה בטוח שברצונך לצאת ", "Inspection Plan", MessageBoxButtons.YesNo);

            if (res == DialogResult.Yes)
            {
                if (_ntlsSite != null) w.CloseQuery();
                //this.Close();
                //this.Hide();
                this.Dispose();

                return true;
            }
            else
            {
                return false;
            }
        }

        public WindowRefreshType DataChange()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public WindowButtonsType GetButtons()
        {
            return LSExtensionWindowLib.WindowButtonsType.windowButtonsNone;
        }

        public void Internationalise()
        {
        }

        public void PreDisplay()
        {
            xmlProcessor = Utils.GetXmlProcessor(sp);

            _ntlsUser = Utils.GetNautilusUser(sp);

            InitializeData();
        }

        public void RestoreSettings(int hKey)
        {
        }

        public bool SaveData()
        {
            return true;
        }

        public void SaveSettings(int hKey)
        {
        }

        public void SetParameters(string parameters)
        {
        }

        public void SetServiceProvider(object serviceProvider)
        {
            sp = serviceProvider as NautilusServiceProvider;
            _ntlsCon = Utils.GetNtlsCon(sp);
        }

        public void SetSite(object site)
        {
            _ntlsSite = (IExtensionWindowSite2)site;
            _ntlsSite.SetWindowInternalName("Inspection_Plan");
            _ntlsSite.SetWindowRegistryName("Inspection_Plan");
            _ntlsSite.SetWindowTitle("Inspection Plan");
        }

        public void Setup()
        {
        }

        public WindowRefreshType ViewRefresh()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public void refresh()
        {
        }

        private InspectionPlanWPF w;
        private void InitializeData()
        {
            w = new InspectionPlanWPF(sp, xmlProcessor, _ntlsCon, _ntlsSite, _ntlsUser);

            elementHost1.Child = w;
        }
    }
}
