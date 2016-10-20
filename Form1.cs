using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace SPSiteAdmin2016
{
    public partial class Form1 : Form
    {
        Dictionary<string, string> _dictPSScriptInWebApp = new Dictionary<string, string>();
        Dictionary<string, string> _dictPSScriptBackup = new Dictionary<string, string>();
        Dictionary<string, string> _dictPSScriptRestore = new Dictionary<string, string>();
        Dictionary<string, string> _dictPSScriptCreate = new Dictionary<string, string>();
        Dictionary<string, string> _dictPSScriptCrossWebApp = new Dictionary<string, string>();

        public const string _SPContentDatabase_Tmp = @"SP_Content_Tmp";
        //site url, database name
        public const string _PSTemplate_SPSiteMove = @"Move-SPSite -Identity {0} -DestinationDatabase {1} -Confirm:$false";
        //url, database name, targetUrl
        public const string _PSTemplate_SPSiteCopy = @"Copy-SPSite -Identity {0} -DestinationDatabase {1} -TargetUrl {2}";
        //site url, folder, file
        public const string _PSTemplate_SPSiteBackup = @"Backup-SPSite {0} -Path ""{1}\{2}""";
        //site url, folder, file, database server, database name
        public const string _PSTemplate_SPSiteRestore = @"Restore-SPSite -Identity {0} -Path ""{1}\{2}"" -Force -confirm:$false";
        //public const string _PSTemplate_RestoreSPSite = @"Restore-SPSite -Identity {0} -Path ""{1}\{2}"" -DatabaseServer {3} -DatabaseName {4} -Force";
        public const string _PSTemplate_SPSiteCreate = @"New-SPSite {0} -OwnerAlias ""{1}"" -Name ""{2}"" -Template ""{3}"" -ContentDatabase {4} -Description ""{5}""";
        // Remove-SPSite -Identity "<URL>"
        public const string _PSTemplate_SPSiteRemove = @"Remove-SPSite -Identity {0} -confirm:$false";
        // Get-SPTimerJob -WebApplication $WebAppUrl job-site-deletion | Start-SPTimerJob
        public const string _PSTemplate_SPSiteDeletionTimerJob = @"Get-SPTimerJob -WebApplication {0} job-site-deletion | Start-SPTimerJob";
        // Dismount-SPContentDatabase -Identity "TempContentDatabaseSource" -Confirm:$false
        public const string _PSTemplate_SPContentDatabaseDismount = @"Dismount-SPContentDatabase -Identity {0} -Confirm:$false";
        // Mount-SPContentDatabase -AssignNewDatabaseId -Name "SP_Content_SPTest80_HRRecords" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
        public const string _PSTemplate_SPContentDatabaseMount = @"Mount-SPContentDatabase -AssignNewDatabaseId -Name {0} -DatabaseServer {1} -WebApplication {2}";
        // $site = Get-SPSite http://sptest2016dev.unitingcare.local/sites/HRRecordsNew
        public const string _PSTemplate_GetSPSite = @"$site = Get-SPSite {0}";
        // $site.RecycleBin.DeleteAll()
        public const string _PSTemplate_SPSite_RecycleBin_DeleteAll = @"$site.RecycleBin.DeleteAll()";
        // $site.Rename("http://sptest2016dev.unitingcare.local/sites/HRRecords")
        public const string _PSTemplate_SPSiteRename = @"$site.Rename(""{0}"")";
        public const string _PSTemplate_SPSiteDispose = @"$site.Dispose()";
        public const string _PSTemplate_RefreshSitesInConfigurationDatabase = @"$site.ContentDatabase.RefreshSitesInConfigurationDatabase()";

        public string _SP_SQL_Instance_Name = @"";

        DataTable dtRestore = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private int GetComboBoxSelectedIndexByValue(ComboBox objComboBox, string strValue)
        {
            int iIndex = -1;

            for (int i = 0; i < objComboBox.Items.Count; i++)
            {
                if (((KeyValuePair<string, string>)objComboBox.Items[i]).Key.Equals(strValue, StringComparison.InvariantCultureIgnoreCase))
                {
                    iIndex = i;
                    break;
                }
            }

            return iIndex;
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            Cursor cursorBackup = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            if (e.TabPage.Name.Equals(tabPageInWebApp.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                initTabInWebApp();
            }
            else if (e.TabPage.Name.Equals(tabPageBackup.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                initTabBackup();
            }
            else if (e.TabPage.Name.Equals(tabPageRestore.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                initTabRestore();
            }
            else if (e.TabPage.Name.Equals(tabPageCreate.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                initTabCreate();
            }
            else if (e.TabPage.Name.Equals(tabPageCrossWebApp.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                initTabCrossWebApp();
            }

            textBoxPSScript.Text = string.Empty;
            Cursor.Current = cursorBackup;
        }

        private void initTabInWebApp()
        {
            string strSPWebAppId = string.Empty;
            string strSelectedValueWebAppMove = SPSiteAdmin2016.Properties.Settings.Default.InWebApp;
            string strSelectedValueContentDBSource = SPSiteAdmin2016.Properties.Settings.Default.InWebAppContentDBSource;
            string strSelectedValueContentDBDest = SPSiteAdmin2016.Properties.Settings.Default.InWebAppContentDBDest;

            comboBoxInWebApp.Items.Clear();
            foreach (SPWebApplication objSPWebApp in SPWebService.ContentService.WebApplications)
            {
                strSPWebAppId = objSPWebApp.Id.ToString();

                comboBoxInWebApp.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
            }

            comboBoxInWebApp.DisplayMember = "Value";
            comboBoxInWebApp.ValueMember = "Key";
            if (string.IsNullOrEmpty(strSelectedValueWebAppMove) == false)
            {
                comboBoxInWebApp.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxInWebApp, strSelectedValueWebAppMove);
                comboBoxInWebAppContentDBSource.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxInWebAppContentDBSource, strSelectedValueContentDBSource);
                comboBoxInWebAppContentDBDest.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxInWebAppContentDBDest, strSelectedValueContentDBDest);
            }
            else
            {
                if (comboBoxInWebApp.Items.Count > 0)
                    comboBoxInWebApp.SelectedIndex = 0;
                if (comboBoxInWebAppContentDBSource.Items.Count > 0)
                    comboBoxInWebAppContentDBSource.SelectedIndex = 0;
                if (comboBoxInWebAppContentDBDest.Items.Count > 0)
                    comboBoxInWebAppContentDBDest.SelectedIndex = 0;
            }
        }

        private void initTabCrossWebApp()
        {
            string strSPWebAppId = string.Empty;
            string strSelectedValueWebAppCrossWebAppSource = SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppSource;
            string strSelectedValueWebAppCrossWebAppDest = SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppDest;
            string strSelectedValueCrossWebAppContentDBSource = SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppContentDBSource;
            string strSelectedValueCrossWebAppContentDBDest = SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppContentDBDest;

            comboBoxCrossWebAppSource.Items.Clear();
            comboBoxCrossWebAppDest.Items.Clear();
            foreach (SPWebApplication objSPWebApp in SPWebService.ContentService.WebApplications)
            {
                strSPWebAppId = objSPWebApp.Id.ToString();

                comboBoxCrossWebAppSource.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
                comboBoxCrossWebAppDest.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
            }

            comboBoxCrossWebAppSource.DisplayMember = "Value";
            comboBoxCrossWebAppSource.ValueMember = "Key";
            if (string.IsNullOrEmpty(strSelectedValueWebAppCrossWebAppSource) == false)
            {
                comboBoxCrossWebAppSource.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxCrossWebAppSource, strSelectedValueWebAppCrossWebAppSource);
                comboBoxCrossWebAppContentDBSource.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxCrossWebAppContentDBSource, strSelectedValueCrossWebAppContentDBSource);
            }
            else
            {
                if (comboBoxCrossWebAppSource.Items.Count > 0)
                    comboBoxCrossWebAppSource.SelectedIndex = 0;
                if (comboBoxInWebAppContentDBSource.Items.Count > 0)
                    comboBoxInWebAppContentDBSource.SelectedIndex = 0;
            }

            comboBoxCrossWebAppDest.DisplayMember = "Value";
            comboBoxCrossWebAppDest.ValueMember = "Key";
            if (string.IsNullOrEmpty(strSelectedValueWebAppCrossWebAppDest) == false)
            {
                comboBoxCrossWebAppDest.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxCrossWebAppDest, strSelectedValueWebAppCrossWebAppDest);
                comboBoxCrossWebAppContentDBDest.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxCrossWebAppContentDBDest, strSelectedValueCrossWebAppContentDBDest);
            }
            else
            {
                if (comboBoxCrossWebAppDest.Items.Count > 0)
                    comboBoxCrossWebAppDest.SelectedIndex = 0;
                if (comboBoxInWebAppContentDBDest.Items.Count > 0)
                    comboBoxInWebAppContentDBDest.SelectedIndex = 0;
            }
        }

        private void initTabBackup()
        {
            string strSPWebAppId = string.Empty;
            string strSelectedValueBackupWebApp = SPSiteAdmin2016.Properties.Settings.Default.WebAppBackup;

            comboBoxBackupWebApp.Items.Clear();
            foreach (SPWebApplication objSPWebApp in SPWebService.ContentService.WebApplications)
            {
                strSPWebAppId = objSPWebApp.Id.ToString();

                comboBoxBackupWebApp.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
            }
            comboBoxBackupWebApp.DisplayMember = "Value";
            comboBoxBackupWebApp.ValueMember = "Key";
            if (string.IsNullOrEmpty(strSelectedValueBackupWebApp) == false)
            {
                comboBoxBackupWebApp.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxBackupWebApp, strSelectedValueBackupWebApp);
            }
            else
            {
                if (comboBoxBackupWebApp.Items.Count > 0)
                    comboBoxBackupWebApp.SelectedIndex = 0;
            }

            if (string.IsNullOrEmpty(SPSiteAdmin2016.Properties.Settings.Default.BackupFolder) == false)
            {
                textBoxBackupFolder.Text = SPSiteAdmin2016.Properties.Settings.Default.BackupFolder;
            }
            else
            {
                textBoxBackupFolder.Text = Environment.CurrentDirectory;
            }

            checkBoxTimestampBackup.Checked = SPSiteAdmin2016.Properties.Settings.Default.TimestampBackup;
        }

        private void initTabRestore()
        {
            string strSPWebAppId = string.Empty;
            string strSelectedValueRestoreWebApp = SPSiteAdmin2016.Properties.Settings.Default.WebAppRestore;

            comboBoxRestoreWebApp.Items.Clear();
            foreach (SPWebApplication objSPWebApp in SPWebService.ContentService.WebApplications)
            {
                strSPWebAppId = objSPWebApp.Id.ToString();

                comboBoxRestoreWebApp.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
            }

            if (dtRestore != null && dtRestore.Columns.Count > 0)
            {
                dtRestore.Rows.Clear();
                dtRestore.Columns.Clear();
            }
            dtRestore.Columns.Add("Selected", typeof(bool));
            dtRestore.Columns.Add("SiteName");
            dtRestore.Columns.Add("SiteUrl");

            dataGridViewRestoreSPSite.DataSource = dtRestore;
            dataGridViewRestoreSPSite.Columns["SiteUrl"].Visible = false;
            dataGridViewRestoreSPSite.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            comboBoxRestoreWebApp.DisplayMember = "Value";
            comboBoxRestoreWebApp.ValueMember = "Key";

            if (string.IsNullOrEmpty(strSelectedValueRestoreWebApp) == false)
            {
                comboBoxRestoreWebApp.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxRestoreWebApp, strSelectedValueRestoreWebApp);
            }
            else
            {
                if (comboBoxRestoreWebApp.Items.Count > 0)
                    comboBoxRestoreWebApp.SelectedIndex = 0;
            }

            if (string.IsNullOrEmpty(SPSiteAdmin2016.Properties.Settings.Default.RestoreFolder) == false)
            {
                textBoxRestoreFolder.Text = SPSiteAdmin2016.Properties.Settings.Default.RestoreFolder;
            }
            else
            {
                textBoxRestoreFolder.Text = Environment.CurrentDirectory;
            }

            PopulateComboBoxFilesRestore(comboBoxFilesRestore);
        }

        private void initTabCreate()
        {
            string strSPWebAppId = string.Empty;
            string strSelectedValueWebAppCreate = SPSiteAdmin2016.Properties.Settings.Default.WebAppCreate;
            string strSelectedValueContentDBCreate = SPSiteAdmin2016.Properties.Settings.Default.ContentDBCreate;

            comboBoxRestoreWebApp.Items.Clear();
            foreach (SPWebApplication objSPWebApp in SPWebService.ContentService.WebApplications)
            {
                strSPWebAppId = objSPWebApp.Id.ToString();

                comboBoxWebAppCreate.Items.Add(new KeyValuePair<string, string>(strSPWebAppId, objSPWebApp.Name));
            }

            comboBoxWebAppCreate.DisplayMember = "Value";
            comboBoxWebAppCreate.ValueMember = "Key";
            if (string.IsNullOrEmpty(strSelectedValueWebAppCreate) == false)
            {
                comboBoxWebAppCreate.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxWebAppCreate, strSelectedValueWebAppCreate);
                comboBoxContentDBCreate.SelectedIndex = GetComboBoxSelectedIndexByValue(comboBoxContentDBCreate, strSelectedValueContentDBCreate);
            }
            else
            {
                if (comboBoxWebAppCreate.Items.Count > 0)
                    comboBoxWebAppCreate.SelectedIndex = 0;
                if (comboBoxContentDBCreate.Items.Count > 0)
                    comboBoxContentDBCreate.SelectedIndex = 0;
            }

            Dictionary<string, string> dictSiteTemplate = new Dictionary<string, string>();
            dictSiteTemplate.Add("STS#0", "Team Site");
            dictSiteTemplate.Add("STS#1", "Blank Site");
            comboBoxTemplateCreate.DataSource = new BindingSource(dictSiteTemplate, null);
            comboBoxTemplateCreate.DisplayMember = "Value";
            comboBoxTemplateCreate.ValueMember = "Key";
            comboBoxTemplateCreate.SelectedIndex = 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Cursor objCurrentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                labelTempDbName.Text = string.Format("Temporary database name: {0}", _SPContentDatabase_Tmp);
                initTabInWebApp();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(@"Exception message: {0}, ex.StackTrace: {1}", ex.Message, ex.StackTrace));
                throw;
            }
            finally
            {
                Cursor.Current = objCurrentCursor;
            }
        }

        private void PopulateContentDBs(Guid guidWebApp, ComboBox comboBoxContentDB)
        {
            if (comboBoxContentDB.Items.Count > 0)
                comboBoxContentDB.Items.Clear();

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[guidWebApp];
            float decSizeWebApp = 0;
            int iSiteCount = 0;
            foreach (SPContentDatabase objSPContentDatabase in objSPWebApp.ContentDatabases)
            {
                if (string.IsNullOrEmpty(_SP_SQL_Instance_Name))
                {
                    _SP_SQL_Instance_Name = objSPContentDatabase.Server;
                }
                comboBoxContentDB.Items.Add(new KeyValuePair<string, string>(objSPContentDatabase.Id.ToString(), objSPContentDatabase.Name));
                decSizeWebApp += objSPContentDatabase.DiskSizeRequired;
                iSiteCount += objSPContentDatabase.CurrentSiteCount;
            }
            textBoxInWebAppSourceWebAppSize.Text = (decSizeWebApp / (1024 * 1024 * 1024)).ToString("#.00");
            textBoxInWebAppSourceSiteCount.Text = iSiteCount.ToString();

            comboBoxContentDB.DisplayMember = "Value";
            comboBoxContentDB.ValueMember = "Key";
        }

        private void PopulateManagedPath(Guid guidWebApp, ComboBox comboBoxManagedPath)
        {
            if (comboBoxManagedPath.Items.Count > 0)
                comboBoxManagedPath.Items.Clear();

            comboBoxManagedPath.Items.Add(new KeyValuePair<string, string>("", ""));

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[guidWebApp];
            foreach (SPPrefix prefix in objSPWebApp.Prefixes)
            {
                if (prefix.PrefixType == SPPrefixType.WildcardInclusion || prefix.PrefixType == SPPrefixType.Wildcard)
                {
                    comboBoxManagedPath.Items.Add(new KeyValuePair<string, string>(prefix.Name, prefix.Name));
                }
            }

            comboBoxManagedPath.DisplayMember = "Value";
            comboBoxManagedPath.ValueMember = "Key";
        }

        private void PopulateSPSites(Guid guidWebApp, Guid guidContentDB, string strManagedPath, ListBox listBoxSPSite)
        {
            if (listBoxSPSite.Items.Count > 0)
                listBoxSPSite.Items.Clear();

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[guidWebApp];
            SPContentDatabase objSPContentDatabase = objSPWebApp.ContentDatabases[guidContentDB];
            foreach (SPSite objSPSite in objSPContentDatabase.Sites)
            {
                try
                {
                    if (objSPSite.Url.EndsWith(@"Office_Viewing_Service_Cache"))
                        continue;

                    if (string.IsNullOrEmpty(strManagedPath) == false)
                    {
                        //string rootSiteCollectionURL = objSPSite.WebApplication.GetResponseUri(SPUrlZone.Default).ToString();
                        if (objSPSite.Url.Contains(string.Format(@"/{0}/", strManagedPath)) == false)
                            continue;
                    }

                    listBoxSPSite.Items.Add(new KeyValuePair<string, string>(objSPSite.Url, string.Format("{0}, {1}", objSPSite.RootWeb.Title, objSPSite.Url)));
                }
                catch (Exception ex)
                {
                    // MessageBox.Show(string.Format("PopulateSPSites(), ex.Message={0}, ex.StackTrace={1}", ex.Message, ex.StackTrace));
                    // MessageBox.Show(string.Format("PopulateSPSites(), ex.Message={0}", ex.Message));
                    //throw;
                }
            }

            listBoxSPSite.DisplayMember = "Value";
            listBoxSPSite.ValueMember = "Key";

            listBoxSPSite.ClearSelected();

            textBoxPSScript.Text = string.Empty;
            _dictPSScriptInWebApp.Clear();
            _dictPSScriptCrossWebApp.Clear();
        }

        private void PopulateSPSites(Guid guidWebApp, ListBox listBoxSPSite)
        {
            if (listBoxSPSite.Items.Count > 0)
                listBoxSPSite.Items.Clear();

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[guidWebApp];
            foreach (SPSite objSPSite in objSPWebApp.Sites)
            {
                if (objSPSite.Url.EndsWith(@"Office_Viewing_Service_Cache"))
                    continue;

                listBoxSPSite.Items.Add(new KeyValuePair<string, string>(objSPSite.Url, string.Format("{0}, {1}", objSPSite.RootWeb.Title, objSPSite.Url)));
            }

            listBoxSPSite.DisplayMember = "Value";
            listBoxSPSite.ValueMember = "Key";

            listBoxSPSite.ClearSelected();

            textBoxPSScript.Text = string.Empty;
            _dictPSScriptBackup.Clear();
        }

        private void PopulateSPSites(Guid guidWebApp, DataGridView dataGridViewSPSite)
        {
            if (dataGridViewSPSite.Columns.Count > 0)
                dataGridViewSPSite.Columns.Clear();
            if (dtRestore.Rows.Count > 0)
                dtRestore.Rows.Clear();

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[guidWebApp];
            foreach (SPSite objSPSite in objSPWebApp.Sites)
            {
                if (objSPSite.Url.EndsWith(@"Office_Viewing_Service_Cache"))
                    continue;

                DataRow newRow = dtRestore.NewRow();
                newRow["Selected"] = 0;
                newRow["SiteName"] = objSPSite.RootWeb.Title;
                newRow["SiteUrl"] = objSPSite.Url;

                dtRestore.Rows.Add(newRow);
            }
            dataGridViewRestoreSPSite.DataSource = null;
            dataGridViewRestoreSPSite.DataSource = dtRestore;

            textBoxPSScript.Text = string.Empty;
            _dictPSScriptRestore.Clear();
        }

        private void ResetStatusListBoxSPSite(ComboBox objComboBoxContentDBSource, ComboBox objComboBoxContentDBDest, ListBox objListBoxMoveSPSiteSource, ListBox objListBoxMoveSPSiteDest)
        {
            bool boolEnable = false;

            if (objComboBoxContentDBSource.SelectedItem == null || objComboBoxContentDBDest.SelectedItem == null)
            {
                ;
            }
            else if (((KeyValuePair<string, string>)objComboBoxContentDBSource.SelectedItem).Key == ((KeyValuePair<string, string>)objComboBoxContentDBDest.SelectedItem).Key
                    && comboBoxInWebAppManagedPath.Items.Count == 1)
            {
                ;
            }
            else
            {
                boolEnable = true;
            }

            objListBoxMoveSPSiteSource.Enabled = boolEnable;
            objListBoxMoveSPSiteDest.Enabled = boolEnable;
        }

        private void buttonInWebAppReset_Click(object sender, EventArgs e)
        {
            buttonClearPSScript_Click(null, null);
            comboBoxInWebAppContentDBSource_SelectedIndexChanged(null, null);
            comboBoxInWebAppContentDBDest_SelectedIndexChanged(null, null);
        }

        private void comboBoxInWebApp_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strDestWebAppGUID = ((KeyValuePair<string, string>)comboBoxInWebApp.SelectedItem).Key;

            PopulateContentDBs(new Guid(strDestWebAppGUID), comboBoxInWebAppContentDBSource);
            PopulateContentDBs(new Guid(strDestWebAppGUID), comboBoxInWebAppContentDBDest);

            SPSiteAdmin2016.Properties.Settings.Default.InWebApp = strDestWebAppGUID;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxInWebAppContentDBSource.Items.Count > 0)
            {
                comboBoxInWebAppContentDBSource.SelectedIndex = 0;
                comboBoxInWebAppContentDBDest.SelectedIndex = 0;
            }

            PopulateManagedPath(new Guid(strDestWebAppGUID), comboBoxInWebAppManagedPath);

            if (comboBoxInWebAppManagedPath.Items.Count > 0)
            {
                comboBoxInWebAppManagedPath.SelectedIndex = 0;
            }
        }

        private bool SetLabelInWebAppPrompt(string strSiteUrl)
        {
            bool boolReturn = false;

            labelInWebAppPrompt.Visible = false;

            if (comboBoxInWebAppContentDBSource.Items.Count < 1)
                return boolReturn;
            if (comboBoxInWebAppContentDBDest.Items.Count < 1)
                return boolReturn;
            if (comboBoxInWebAppContentDBSource.SelectedItem == null)
                return boolReturn;
            if (comboBoxInWebAppContentDBDest.SelectedItem == null)
                return boolReturn;

            //string strKeyContentDBSource = ((KeyValuePair<string, string>)comboBoxInWebAppContentDBSource.SelectedItem).Key;
            //string strKeyContentDBDest = ((KeyValuePair<string, string>)comboBoxInWebAppContentDBDest.SelectedItem).Key;

            if (comboBoxInWebAppContentDBSource.Text.Equals(comboBoxInWebAppContentDBDest.Text, StringComparison.InvariantCultureIgnoreCase))
            {
                if (comboBoxInWebAppManagedPath.Items.Count == 1)
                {
                    boolReturn = true;
                }
            }

            if (string.IsNullOrEmpty(strSiteUrl) == false && boolReturn == false)
            {
                string strManagedPathSource = GetSitePathName(strSiteUrl);
                string strManagedpathDest = ((KeyValuePair<string, string>)comboBoxInWebAppManagedPath.SelectedItem).Value;
                if (strManagedpathDest.Equals(strManagedPathSource, StringComparison.InvariantCultureIgnoreCase))
                {
                    boolReturn = true;
                }
            }

            labelInWebAppPrompt.Visible = boolReturn;
            return boolReturn;
        }

        private void SetLabelCrossWebAppPrompt()
        {
            labelCrossWebAppPrompt.Visible = false;

            if (comboBoxCrossWebAppContentDBSource.Items.Count < 1)
                return;
            if (comboBoxCrossWebAppContentDBDest.Items.Count < 1)
                return;
            if (comboBoxCrossWebAppContentDBSource.SelectedItem == null)
                return;
            if (comboBoxCrossWebAppContentDBDest.SelectedItem == null)
                return;

            string strKeyContentDBSource = ((KeyValuePair<string, string>)comboBoxCrossWebAppContentDBSource.SelectedItem).Key;
            string strKeyContentDBDest = ((KeyValuePair<string, string>)comboBoxCrossWebAppContentDBDest.SelectedItem).Key;

            if (strKeyContentDBSource.Equals(strKeyContentDBDest, StringComparison.InvariantCultureIgnoreCase))
                labelCrossWebAppPrompt.Visible = true;
        }

        private void comboBoxInWebAppContentDBSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKeyWebApp = ((KeyValuePair<string, string>)comboBoxInWebApp.SelectedItem).Key;
            string strKeyContentDB = ((KeyValuePair<string, string>)comboBoxInWebAppContentDBSource.SelectedItem).Key;

            PopulateSPSites(new Guid(strKeyWebApp),
                new Guid(strKeyContentDB),
                @"",
                listBoxInWebAppSPSiteSource);

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[new Guid(strKeyWebApp)];
            SPContentDatabase objSPContentDatabase = objSPWebApp.ContentDatabases[new Guid(strKeyContentDB)];
            textBoxInWebAppSourceDBSize.Text = (((float)objSPContentDatabase.DiskSizeRequired) / (1024 * 1024 * 1024)).ToString("#.00");
            textBoxInWebAppSourceDBSiteCount.Text = objSPContentDatabase.CurrentSiteCount.ToString();

            SPSiteAdmin2016.Properties.Settings.Default.InWebAppContentDBSource = strKeyContentDB;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            ResetStatusListBoxSPSite(comboBoxInWebAppContentDBSource, comboBoxInWebAppContentDBDest, listBoxInWebAppSPSiteSource, listBoxInWebAppSPSiteDest);
            SetLabelInWebAppPrompt(string.Empty);
        }

        private void comboBoxInWebAppContentDBDest_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKeyWebApp = ((KeyValuePair<string, string>)comboBoxInWebApp.SelectedItem).Key;
            string strKeyContentDB = ((KeyValuePair<string, string>)comboBoxInWebAppContentDBDest.SelectedItem).Key;
            string strManagedPath = string.Empty;

            if (_dictPSScriptInWebApp.Count > 0)
            {
                // RefreshPSScriptTextBoxInWebApp();
                buttonClearPSScript_Click(null, null);
                comboBoxInWebAppContentDBSource_SelectedIndexChanged(null, null);
            }

            if (comboBoxInWebAppManagedPath.SelectedItem != null)
            {
                strManagedPath = ((KeyValuePair<string, string>)comboBoxInWebAppManagedPath.SelectedItem).Value;
            }
            PopulateSPSites(new Guid(strKeyWebApp),
                new Guid(strKeyContentDB),
                strManagedPath,
                listBoxInWebAppSPSiteDest);

            SPWebApplication objSPWebApp = SPWebService.ContentService.WebApplications[new Guid(strKeyWebApp)];
            SPContentDatabase objSPContentDatabase = objSPWebApp.ContentDatabases[new Guid(strKeyContentDB)];
            textBoxInWebAppDestDBSize.Text = (((float)objSPContentDatabase.DiskSizeRequired) / (1024 * 1024 * 1024)).ToString("#.00");
            textBoxInWebAppDestDBSiteCount.Text = objSPContentDatabase.CurrentSiteCount.ToString();

            SPSiteAdmin2016.Properties.Settings.Default.InWebAppContentDBDest = strKeyContentDB;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            ResetStatusListBoxSPSite(comboBoxInWebAppContentDBSource, comboBoxInWebAppContentDBDest, listBoxInWebAppSPSiteSource, listBoxInWebAppSPSiteDest);
            SetLabelInWebAppPrompt(string.Empty);
        }

        private void comboBoxInWebAppManagedPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxInWebAppContentDBDest_SelectedIndexChanged(sender, e);
        }

        private string getWebAppUrlByGUID(string strWebAppGUID)
        {
            string strWebAppUrl = string.Empty;
            SPWebApplication objSourceSPWebApp = SPWebService.ContentService.WebApplications[new Guid(strWebAppGUID)];
            foreach (SPAlternateUrl altUrl in objSourceSPWebApp.AlternateUrls)
            {
                if (altUrl.UrlZone == SPUrlZone.Default)
                {
                    strWebAppUrl = altUrl.Uri.ToString();
                    break;
                }
            }

            return strWebAppUrl;
        }

        private void RefreshPSScriptTextBoxInWebApp()
        {
            string strLine = string.Empty;
            textBoxPSScript.Text = string.Empty;

            string strWebAppGUID = ((KeyValuePair<string, string>)comboBoxInWebApp.SelectedItem).Key;
            string strWebAppUrl = getWebAppUrlByGUID(strWebAppGUID);
            string strSiteUrlDest = string.Empty;

            foreach (string strSiteUrlSource in _dictPSScriptInWebApp.Keys)
            {
                // Move-SPSite http://sharepoint/sites/moveme -DestinationDatabase WSS_Content2 
                if (radioButtonInWebAppMove.Checked)
                {
                    if (comboBoxInWebAppContentDBSource.Text.Equals(comboBoxInWebAppContentDBDest.Text, StringComparison.InvariantCultureIgnoreCase) == false)
                    {
                        strLine = string.Format(_PSTemplate_SPSiteMove, strSiteUrlSource, comboBoxInWebAppContentDBDest.Text);
                        textBoxPSScript.Text += strLine + Environment.NewLine;
                    }
                    else
                    {
                        string strManagedPathSource = GetSitePathName(strSiteUrlSource);
                        string strManagedpathDest = ((KeyValuePair<string, string>)comboBoxInWebAppManagedPath.SelectedItem).Value;
                        strSiteUrlDest = strSiteUrlSource.Replace(strManagedPathSource, strManagedpathDest);
                        string strContentDB = ((KeyValuePair<string, string>)comboBoxInWebAppContentDBDest.SelectedItem).Value;

                        changeSiteManagedPath(strWebAppUrl, strContentDB, strSiteUrlSource, strSiteUrlDest);
                    }
                }
                else
                {
                    strSiteUrlDest = strSiteUrlSource + "New";

                    strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrl);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPSiteCopy, strSiteUrlSource, comboBoxInWebAppContentDBDest.Text, strSiteUrlDest);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPSiteMove, strSiteUrlDest, comboBoxInWebAppContentDBSource.Text);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;
                }
            }
        }

        private void listBoxInWebApp_DragDrop(object sender, DragEventArgs e, ListBox objListBoxSource, ListBox objListBoxDest)
        {
            string strKey_SiteUrl = string.Empty;
            string strValue = string.Empty;
            ListBox.SelectedObjectCollection objItems = objListBoxSource.SelectedItems;
            int iCurrent = int.MinValue;
            bool boolChangeFlag = false;

            if (objListBoxSource.SelectedItem == null)
                return;

            strKey_SiteUrl = ((KeyValuePair<string, string>)objListBoxSource.SelectedItem).Key;
            strValue = ((KeyValuePair<string, string>)objListBoxSource.SelectedItem).Value;
            if (objListBoxSource == listBoxInWebAppSPSiteSource)
            {
                if (_dictPSScriptInWebApp.ContainsKey(strKey_SiteUrl) == false)
                {
                    _dictPSScriptInWebApp.Add(strKey_SiteUrl, strValue);
                    boolChangeFlag = true;
                }
            }
            else
            {
                if (_dictPSScriptInWebApp.ContainsKey(strKey_SiteUrl))
                {
                    _dictPSScriptInWebApp.Remove(strKey_SiteUrl);
                    boolChangeFlag = true;
                }
            }

            if (boolChangeFlag)
            {
                bool boolReturn = SetLabelInWebAppPrompt(strKey_SiteUrl);
                if (boolReturn)
                {
                    _dictPSScriptInWebApp.Remove(strKey_SiteUrl);
                    return;
                }

                foreach (var item in objItems)
                {
                    iCurrent = objListBoxDest.Items.Add(item);
                }
                while (objListBoxSource.SelectedItems.Count > 0)
                {
                    objListBoxSource.Items.Remove(objListBoxSource.SelectedItems[0]);
                }
                RefreshPSScriptTextBoxInWebApp();
            }
        }

        private void renameSiteUrl(string strSiteUrlSource, string strSiteUrlDest)
        {
            string strLine = string.Empty;

            //$site = Get-SPSite http://portal.odfbdemo.com/sites/oldpath 
            //$uri = New-Object System.Uri("http://portal.odfbdemo.com/sites/shinynewpath")
            //$site.Rename($uri)
            //((Get-SPSite http://portal.odfbdemo.com/sites/shinynewpath).contentdatabase).RefreshSitesInConfigurationDatabase

            strLine = string.Format(_PSTemplate_GetSPSite, strSiteUrlSource);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            textBoxPSScript.Text += _PSTemplate_SPSite_RecycleBin_DeleteAll + Environment.NewLine;

            strLine = string.Format(_PSTemplate_SPSiteRename, strSiteUrlDest);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            textBoxPSScript.Text += _PSTemplate_SPSiteDispose + Environment.NewLine;

            strLine = string.Format(_PSTemplate_GetSPSite, strSiteUrlDest);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            textBoxPSScript.Text += _PSTemplate_RefreshSitesInConfigurationDatabase + Environment.NewLine;

            textBoxPSScript.Text += _PSTemplate_SPSiteDispose + Environment.NewLine;
        }

        private void changeSiteManagedPath(string strWebAppUrl, string strContentDatabaseName, string strSiteUrlSource, string strSiteUrlDest)
        {
            string strLine = string.Empty;

            // Mount-SPContentDatabase -AssignNewDatabaseId -Name "TempContentDatabaseSource" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
            strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrl);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            //Copy-SPSite http://$webAppSource$envSuffix/sites/$siteNameSource -DestinationDatabase $databaseNameTmp -TargetUrl http://$webAppSource$envSuffix/sites/$siteNameDest
            strLine = string.Format(_PSTemplate_SPSiteCopy, strSiteUrlSource, _SPContentDatabase_Tmp, strSiteUrlDest);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            strLine = string.Format(_PSTemplate_SPSiteRemove, strSiteUrlSource);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            strLine = string.Format(_PSTemplate_SPSiteDeletionTimerJob, strWebAppUrl);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            //Move-SPSite http://$webAppSource$envSuffix/sites/$siteNameDest -DestinationDatabase $databaseNameSource -Confirm:$false
            strLine = string.Format(_PSTemplate_SPSiteMove, strSiteUrlDest, strContentDatabaseName);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            strLine = string.Format(_PSTemplate_SPSiteDeletionTimerJob, strWebAppUrl);
            textBoxPSScript.Text += strLine + Environment.NewLine;

            //Dismount-SPContentDatabase - Identity $databaseNameTmp - Confirm:$false
            strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
            textBoxPSScript.Text += strLine + Environment.NewLine;
        }

        private void RefreshPSScriptTextBoxCrossWebApp()
        {
            //string strSiteUrl = string.Empty;
            string strFileFullPath = string.Empty;
            string strLine = string.Empty;
            textBoxPSScript.Text = string.Empty;

            string strManagedPathDest = ((KeyValuePair<string, string>)comboBoxCrossWebAppManagedPath.SelectedItem).Key;
            string strContentDB = ((KeyValuePair<string, string>)comboBoxCrossWebAppContentDBDest.SelectedItem).Value;

            string strSourceWebAppGUID = ((KeyValuePair<string, string>)comboBoxCrossWebAppSource.SelectedItem).Key;
            string strWebAppUrlSource = getWebAppUrlByGUID(strSourceWebAppGUID);

            string strDestWebAppGUID = ((KeyValuePair<string, string>)comboBoxCrossWebAppDest.SelectedItem).Key;
            string strWebAppUrlDest = getWebAppUrlByGUID(strDestWebAppGUID);

            string strSiteUrlTmp = string.Empty;
            string strSiteUrlDest = string.Empty;
            string strManagedPathSource = string.Empty;

            //for (int i = 0; i < listBoxCrossWebAppSPSiteSource.SelectedItems.Count; i++)
            foreach (string strSourceSiteUrl in _dictPSScriptCrossWebApp.Keys)
            {
                strSiteUrlDest = strSourceSiteUrl.Replace(strWebAppUrlSource, strWebAppUrlDest);

                if (radioButtonCrossWebAppMove.Checked)
                {
                    // Mount-SPContentDatabase -AssignNewDatabaseId -Name "TempContentDatabaseSource" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
                    strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrlSource);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Move-SPSite http://$webAppSource$envSuffix/sites/$siteNameDest -DestinationDatabase $databaseNameSource -Confirm:$false
                    strLine = string.Format(_PSTemplate_SPSiteMove, strSourceSiteUrl, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPSiteDeletionTimerJob, strWebAppUrlSource);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Dismount-SPContentDatabase - Identity $databaseNameTmp - Confirm:$false
                    strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    // Mount-SPContentDatabase -AssignNewDatabaseId -Name "TempContentDatabaseSource" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
                    strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrlDest);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Move-SPSite http://$webAppSource$envSuffix/sites/$siteNameDest -DestinationDatabase $databaseNameSource -Confirm:$false
                    strLine = string.Format(_PSTemplate_SPSiteMove, strSiteUrlDest, comboBoxCrossWebAppContentDBDest.Text);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPSiteDeletionTimerJob, strWebAppUrlDest);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Dismount-SPContentDatabase - Identity $databaseNameTmp - Confirm:$false
                    strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;
                }
                else
                {
                    // Mount-SPContentDatabase -AssignNewDatabaseId -Name "TempContentDatabaseSource" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
                    strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrlSource);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Copy-SPSite http://$webAppSource$envSuffix/sites/$siteNameSource -DestinationDatabase $databaseNameTmp -TargetUrl http://$webAppSource$envSuffix/sites/$siteNameDest
                    strLine = string.Format(_PSTemplate_SPSiteCopy, strSourceSiteUrl, _SPContentDatabase_Tmp, strSourceSiteUrl + "New");
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Dismount-SPContentDatabase - Identity $databaseNameTmp - Confirm:$false
                    strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Mount-SPContentDatabase -AssignNewDatabaseId -Name "TempContentDatabaseSource" -DatabaseServer "sp2016dbDEV" -WebApplication "http://SPTest2016DEV.unitingcare.local"
                    strLine = string.Format(_PSTemplate_SPContentDatabaseMount, _SPContentDatabase_Tmp, _SP_SQL_Instance_Name, strWebAppUrlDest);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Move-SPSite http://$webAppSource$envSuffix/sites/$siteNameDest -DestinationDatabase $databaseNameSource -Confirm:$false
                    strLine = string.Format(_PSTemplate_SPSiteMove, strSiteUrlDest + "New", comboBoxCrossWebAppContentDBDest.Text);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    strLine = string.Format(_PSTemplate_SPSiteDeletionTimerJob, strWebAppUrlDest);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    //Dismount-SPContentDatabase - Identity $databaseNameTmp - Confirm:$false
                    strLine = string.Format(_PSTemplate_SPContentDatabaseDismount, _SPContentDatabase_Tmp);
                    textBoxPSScript.Text += strLine + Environment.NewLine;

                    renameSiteUrl(strSiteUrlDest + "New", strSiteUrlDest);
                }
                strManagedPathSource = GetSitePathName(strSourceSiteUrl);
                if (strManagedPathSource.Equals(strManagedPathDest) != true)
                {
                    strSiteUrlTmp = strSiteUrlDest;
                    strSiteUrlDest = strSiteUrlDest.Replace(strManagedPathSource, strManagedPathDest);

                    changeSiteManagedPath(strWebAppUrlDest, strContentDB, strSiteUrlTmp, strSiteUrlDest);
                }
            }
        }

        private void listBoxCrossWebApp_DragDrop(object sender, DragEventArgs e, ListBox objListBoxSource, ListBox objListBoxDest)
        {
            string strKey = string.Empty;
            string strValue = string.Empty;
            ListBox.SelectedObjectCollection objItems = objListBoxSource.SelectedItems;
            int iCurrent = int.MinValue;
            bool boolChangeFlag = false;

            if (objListBoxSource.SelectedItem == null)
                return;

            strKey = ((KeyValuePair<string, string>)objListBoxSource.SelectedItem).Key;
            strValue = ((KeyValuePair<string, string>)objListBoxSource.SelectedItem).Value;
            if (objListBoxSource == listBoxCrossWebAppSPSiteSource)
            {
                if (_dictPSScriptCrossWebApp.ContainsKey(strKey) == false)
                {
                    _dictPSScriptCrossWebApp.Add(strKey, strValue);
                    boolChangeFlag = true;
                }
            }
            else
            {
                if (_dictPSScriptCrossWebApp.ContainsKey(strKey))
                {
                    _dictPSScriptCrossWebApp.Remove(strKey);
                    boolChangeFlag = true;
                }
            }

            if (boolChangeFlag)
            {
                foreach (var item in objItems)
                {
                    iCurrent = objListBoxDest.Items.Add(item);
                }
                while (objListBoxSource.SelectedItems.Count > 0)
                {
                    objListBoxSource.Items.Remove(objListBoxSource.SelectedItems[0]);
                }
                RefreshPSScriptTextBoxCrossWebApp();
            }
        }

        private void listBoxInWebAppSPSiteDest_DragDrop(object sender, DragEventArgs e)
        {
            listBoxInWebApp_DragDrop(sender, e, listBoxInWebAppSPSiteSource, listBoxInWebAppSPSiteDest);
        }

        private void listBoxInWebAppSPSiteSource_DragDrop(object sender, DragEventArgs e)
        {
            listBoxInWebApp_DragDrop(sender, e, listBoxInWebAppSPSiteDest, listBoxInWebAppSPSiteSource);
        }

        private void listBoxInWebAppSPSiteSource_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxInWebAppSPSiteSource.SelectedItem == null)
                return;
            listBoxInWebAppSPSiteSource.DoDragDrop(listBoxInWebAppSPSiteSource.SelectedItem, DragDropEffects.Move);
        }

        private void listBoxInWebAppSPSiteDest_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxInWebAppSPSiteDest.SelectedItem == null) return;
            listBoxInWebAppSPSiteDest.DoDragDrop(listBoxInWebAppSPSiteDest.SelectedItem, DragDropEffects.Move);
        }

        private void listBoxInWebAppSPSiteSource_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void listBoxInWebAppSPSiteDest_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {

        }

        private void buttonInWebAppAll_Click(object sender, EventArgs e)
        {
            for (int i = listBoxInWebAppSPSiteSource.Items.Count - 1; i >= 0; i--)
            {
                listBoxInWebAppSPSiteSource.SelectedIndex = i;

                listBoxInWebApp_DragDrop(null, null, listBoxInWebAppSPSiteSource, listBoxInWebAppSPSiteDest);
            }
        }

        private string GetFolderPath()
        {
            string strFolderPath = string.Empty;

            folderBrowserDialog1.ShowNewFolderButton = true;
            strFolderPath = textBoxBackupFolder.Text.Trim();
            if (string.IsNullOrEmpty(strFolderPath) == false)
            {
                if (System.IO.Directory.Exists(strFolderPath))
                {
                    folderBrowserDialog1.SelectedPath = strFolderPath;
                }
            }
            DialogResult objDialogResult = folderBrowserDialog1.ShowDialog();
            if (objDialogResult == System.Windows.Forms.DialogResult.OK)
            {
                strFolderPath = folderBrowserDialog1.SelectedPath;
            }

            return strFolderPath;
        }

        private void buttonBrowseBackup_Click(object sender, EventArgs e)
        {
            string strFolderPath = string.Empty;
            strFolderPath = GetFolderPath();
            if (string.IsNullOrEmpty(strFolderPath) == false)
            {
                SPSiteAdmin2016.Properties.Settings.Default.BackupFolder = strFolderPath;
                SPSiteAdmin2016.Properties.Settings.Default.Save();
                textBoxBackupFolder.Text = strFolderPath;
            }
        }

        private void comboBoxBackupWebApp_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKeyWebApp = ((KeyValuePair<string, string>)comboBoxBackupWebApp.SelectedItem).Key;

            PopulateSPSites(new Guid(strKeyWebApp),
                listBoxBackupSPSite);

            SPSiteAdmin2016.Properties.Settings.Default.WebAppBackup = strKeyWebApp;
            SPSiteAdmin2016.Properties.Settings.Default.Save();
        }

        private string GetFileName(string strSiteUrl)
        {
            string strFileName = string.Empty;
            int iPos = 0;

            iPos = strSiteUrl.IndexOf(@"//");
            strFileName = strSiteUrl.Substring(iPos + 2);
            strFileName = strFileName.Replace(@"/", @".");
            strFileName = strFileName.Replace(@":", @".");
            if (checkBoxTimestampBackup.Checked)
            {
                strFileName += @"." + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            }
            strFileName += ".bak";

            return strFileName;
        }

        private string GetSitePathName(string strSiteUrl)
        {
            string strSitePathName = string.Empty;
            int iPosFirst = int.MinValue;
            int iPosSecond = int.MinValue;

            using (SPSite site = new SPSite(strSiteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    iPosFirst = web.ServerRelativeUrl.IndexOf('/');
                    if (iPosFirst >= 0)
                    {
                        iPosSecond = web.ServerRelativeUrl.IndexOf('/', iPosFirst + 1);
                        if (iPosSecond > 0)
                        {
                            strSitePathName = web.ServerRelativeUrl.Substring(iPosFirst + 1, iPosSecond - iPosFirst - 1);
                        }
                    }
                }
            }
            return strSitePathName;
        }

        private void textBoxFolderBackup_TextChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxBackup();
        }

        private void RefreshPSScriptTextBoxBackup()
        {
            string strKey = string.Empty;
            string strLine = string.Empty;
            string strFileFullPath = string.Empty;
            textBoxPSScript.Text = string.Empty;

            for (int i = 0; i < listBoxBackupSPSite.SelectedItems.Count; i++)
            {
                strKey = ((KeyValuePair<string, string>)listBoxBackupSPSite.SelectedItems[i]).Key;
                strFileFullPath = GetFileName(strKey);
                strLine = string.Format(_PSTemplate_SPSiteBackup, strKey, textBoxBackupFolder.Text, strFileFullPath);
                textBoxPSScript.Text += strLine + Environment.NewLine;
            }
        }

        private void listBoxBackupSPSite_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxBackup();
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            listBoxBackupSPSite.SelectedItems.Clear();
            for (int i = listBoxBackupSPSite.Items.Count - 1; i >= 0; i--)
            {
                listBoxBackupSPSite.SelectedItems.Add(listBoxBackupSPSite.Items[i]);
            }
            RefreshPSScriptTextBoxBackup();
        }

        private void dataGridViewRestoreSPSite_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
            ch1 = (DataGridViewCheckBoxCell)dataGridViewRestoreSPSite.Rows[dataGridViewRestoreSPSite.CurrentRow.Index].Cells[0];

            if (ch1.Value == null)
                ch1.Value = false;
            switch (ch1.Value.ToString())
            {
                case "True":
                    ch1.Value = false;
                    break;
                case "False":
                    ch1.Value = true;
                    break;
            }
            RefreshPSScriptTextBoxRestore();
            //MessageBox.Show(ch1.Value.ToString());
        }

        private void comboBoxRestoreWebApp_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKeyWebApp = ((KeyValuePair<string, string>)comboBoxRestoreWebApp.SelectedItem).Key;

            PopulateSPSites(new Guid(strKeyWebApp), dataGridViewRestoreSPSite);

            SPSiteAdmin2016.Properties.Settings.Default.WebAppRestore = strKeyWebApp;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            RefreshPSScriptTextBoxRestore();
        }

        private void buttonBrowseRestore_Click(object sender, EventArgs e)
        {
            string strFolderName = string.Empty;

            folderBrowserDialog1.ShowNewFolderButton = true;
            strFolderName = textBoxRestoreFolder.Text.Trim();
            if (string.IsNullOrEmpty(strFolderName) == false)
            {
                if (System.IO.Directory.Exists(strFolderName))
                {
                    folderBrowserDialog1.SelectedPath = strFolderName;
                }
            }
            DialogResult objDialogResult = folderBrowserDialog1.ShowDialog();
            if (objDialogResult == System.Windows.Forms.DialogResult.OK)
            {
                strFolderName = folderBrowserDialog1.SelectedPath;
                SPSiteAdmin2016.Properties.Settings.Default.RestoreFolder = strFolderName;
                SPSiteAdmin2016.Properties.Settings.Default.Save();
                textBoxRestoreFolder.Text = strFolderName;
            }
        }

        private void textBoxFolderRestore_TextChanged(object sender, EventArgs e)
        {
            PopulateComboBoxFilesRestore(comboBoxFilesRestore);
        }

        private void PopulateComboBoxFilesRestore(ComboBox objComboBox)
        {
            if (objComboBox.Items.Count > 0)
                objComboBox.Items.Clear();

            string targetDirectory = textBoxRestoreFolder.Text;
            if (Directory.Exists(targetDirectory) == false)
                return;

            int iPos = targetDirectory.Length;
            string[] fileEntries = Directory.GetFiles(targetDirectory, @"*.bak", SearchOption.TopDirectoryOnly);

            foreach (string strEntry in fileEntries)
                objComboBox.Items.Add(strEntry.Substring(iPos + 1));

            if (objComboBox.Items.Count > 0)
                objComboBox.SelectedIndex = 0;

            if (objComboBox.SelectedItem == null)
                dataGridViewRestoreSPSite.Enabled = false;
            else
                dataGridViewRestoreSPSite.Enabled = true;
        }

        private void RefreshPSScriptTextBoxRestore()
        {
            string strSiteUrl = string.Empty;
            string strLine = string.Empty;
            string strFileFullPath = string.Empty;
            textBoxPSScript.Text = string.Empty;

            for (int i = 0; i < dataGridViewRestoreSPSite.Rows.Count; i++)
            {
                DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                ch1 = (DataGridViewCheckBoxCell)dataGridViewRestoreSPSite.Rows[i].Cells[0];
                if (ch1.Value == null)
                    ch1.Value = false;

                if ((bool)ch1.Value == true)
                {
                    strSiteUrl = (string)dataGridViewRestoreSPSite.Rows[i].Cells[2].Value;
                    strFileFullPath = comboBoxFilesRestore.SelectedItem.ToString();
                    //site url, folder, file
                    strLine = string.Format(_PSTemplate_SPSiteRestore, strSiteUrl, textBoxRestoreFolder.Text, strFileFullPath);
                    textBoxPSScript.Text += strLine + Environment.NewLine;
                }
            }
        }

        private void comboBoxFilesRestore_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxRestore();
        }

        private void comboBoxWebAppCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKey = ((KeyValuePair<string, string>)comboBoxWebAppCreate.SelectedItem).Key;

            PopulateContentDBs(new Guid(strKey),
                comboBoxContentDBCreate);

            SPSiteAdmin2016.Properties.Settings.Default.WebAppCreate = strKey;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxContentDBCreate.Items.Count > 0)
            {
                comboBoxContentDBCreate.SelectedIndex = 0;
            }

            PopulateManagedPath(new Guid(strKey), comboBoxCreateManagedPath);

            SPSiteAdmin2016.Properties.Settings.Default.CreateManagedPath = strKey;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxCreateManagedPath.Items.Count > 0)
            {
                comboBoxCreateManagedPath.SelectedIndex = 0;
            }
        }

        private void PopulateSPSitesCreate()
        {
            string strKeyWebApp = ((KeyValuePair<string, string>)comboBoxWebAppCreate.SelectedItem).Key;
            string strKeyContentDB = ((KeyValuePair<string, string>)comboBoxContentDBCreate.SelectedItem).Key;
            string strManagedPath = string.Empty;

            if (comboBoxCreateManagedPath.SelectedItem != null)
            {
                strManagedPath = ((KeyValuePair<string, string>)comboBoxCreateManagedPath.SelectedItem).Value;
            }
            PopulateSPSites(new Guid(strKeyWebApp),
                new Guid(strKeyContentDB),
                strManagedPath,
                listBoxSPSiteCreate);
        }

        private void comboBoxContentDBCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateSPSitesCreate();

            //string strKeyContentDB = ((KeyValuePair<string, string>)comboBoxContentDBCreate.SelectedItem).Key;
            //SPSiteAdmin2016.Properties.Settings.Default.ContentDBCreate = strKeyContentDB;
            //SPSiteAdmin2016.Properties.Settings.Default.Save();

            RefreshPSScriptTextBoxCreate();
        }

        private void buttonCreateRefresh_Click(object sender, EventArgs e)
        {
            comboBoxContentDBCreate_SelectedIndexChanged(sender, e);
        }

        private void RefreshPSScriptTextBoxCreate()
        {
            string strSitePathName = textBoxCreatePath.Text.Trim();
            string strSiteName = textBoxCreateName.Text.Trim();

            if (comboBoxTemplateCreate.SelectedItem == null)
                return;

            if (string.IsNullOrEmpty(strSitePathName) || string.IsNullOrEmpty(strSiteName))
                return;

            string strKeyWebAppGUID = ((KeyValuePair<string, string>)comboBoxWebAppCreate.SelectedItem).Key;
            string strWebAppUrl = getWebAppUrlByGUID(strKeyWebAppGUID);

            string strManagedPath = ((KeyValuePair<string, string>)comboBoxCreateManagedPath.SelectedItem).Key;
            string strContentDB = ((KeyValuePair<string, string>)comboBoxContentDBCreate.SelectedItem).Value;

            string strSiteFullPath = string.Format(@"{0}{1}/{2}", strWebAppUrl, strManagedPath, strSitePathName);
            string strSiteTemplate = ((KeyValuePair<string, string>)comboBoxTemplateCreate.SelectedItem).Key;

            string strSiteDescription = textBoxCreateDescription.Text.Trim();
            string strSiteCollectionOwnerLogin = string.Format(@"{0}\{1}", Environment.UserDomainName, Environment.UserName);

            textBoxPSScript.Text = string.Empty;

            // @"New-SPSite {0} -OwnerAlias ""{1}"" –Language 1033 -Name ""{2}"" -Template ""{3}"" -ContentDatabase {4} -Description {5}";
            textBoxPSScript.Text = string.Format(_PSTemplate_SPSiteCreate, strSiteFullPath, strSiteCollectionOwnerLogin,
                strSiteName, strSiteTemplate, strContentDB, strSiteDescription);
        }

        private void comboBoxManagedPathCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxCreate();
        }

        private void textBoxCreatePath_TextChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxCreate();
        }

        private void textBoxCreateName_TextChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxCreate();
        }

        private void comboBoxTemplateCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxCreate();
        }

        private void textBoxCreateDescription_TextChanged(object sender, EventArgs e)
        {
            RefreshPSScriptTextBoxCreate();
        }

        private void buttonClearPSScript_Click(object sender, EventArgs e)
        {
            textBoxPSScript.Text = string.Empty;

            if (_dictPSScriptInWebApp.Count > 0) _dictPSScriptInWebApp.Clear();
            if (_dictPSScriptBackup.Count > 0) _dictPSScriptBackup.Clear();
            if (_dictPSScriptRestore.Count > 0) _dictPSScriptRestore.Clear();
            if (_dictPSScriptCreate.Count > 0) _dictPSScriptCreate.Clear();
            if (_dictPSScriptCrossWebApp.Count > 0) _dictPSScriptCrossWebApp.Clear();
        }
        private void checkBoxTimestampBackup_CheckedChanged(object sender, EventArgs e)
        {
            SPSiteAdmin2016.Properties.Settings.Default.TimestampBackup = checkBoxTimestampBackup.Checked;
        }

        private void buttonRestoreSelectAll_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxCrossWebAppSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKey = ((KeyValuePair<string, string>)comboBoxCrossWebAppSource.SelectedItem).Key;

            PopulateContentDBs(new Guid(strKey), comboBoxCrossWebAppContentDBSource);

            SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppSource = strKey;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxInWebAppContentDBSource.Items.Count > 0)
            {
                comboBoxInWebAppContentDBSource.SelectedIndex = 0;
                comboBoxInWebAppContentDBDest.SelectedIndex = 0;
            }
        }

        private void comboBoxCrossWebAppDest_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strDestWebAppGUID = ((KeyValuePair<string, string>)comboBoxCrossWebAppDest.SelectedItem).Key;

            PopulateContentDBs(new Guid(strDestWebAppGUID), comboBoxCrossWebAppContentDBDest);

            SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppDest = strDestWebAppGUID;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxCrossWebAppContentDBDest.Items.Count > 0)
            {
                comboBoxCrossWebAppContentDBDest.SelectedIndex = 0;
            }

            PopulateManagedPath(new Guid(strDestWebAppGUID), comboBoxCrossWebAppManagedPath);

            SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppDest = strDestWebAppGUID;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            if (comboBoxCrossWebAppManagedPath.Items.Count > 0)
            {
                comboBoxCrossWebAppManagedPath.SelectedIndex = 0;
            }
        }

        private void comboBoxCrossWebAppContentDBDest_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxCrossWebAppDest.SelectedItem == null)
                return;
            if (comboBoxCrossWebAppContentDBDest.SelectedItem == null)
                return;

            RefreshPSScriptTextBoxCrossWebApp();

            string strKeyCrossWebAppDest = ((KeyValuePair<string, string>)comboBoxCrossWebAppDest.SelectedItem).Key;
            string strKeyCrossWebAppContentDBDest = ((KeyValuePair<string, string>)comboBoxCrossWebAppContentDBDest.SelectedItem).Key;
            string strManagedPath = string.Empty;

            if (comboBoxCreateManagedPath.SelectedItem != null)
            {
                strManagedPath = ((KeyValuePair<string, string>)comboBoxCrossWebAppManagedPath.SelectedItem).Value;
            }
            PopulateSPSites(new Guid(strKeyCrossWebAppDest),
                new Guid(strKeyCrossWebAppContentDBDest),
                strManagedPath,
                listBoxCrossWebAppSPSiteDest);

            SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppContentDBDest = strKeyCrossWebAppContentDBDest;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            ResetStatusListBoxSPSite(comboBoxCrossWebAppContentDBSource, comboBoxCrossWebAppContentDBDest, listBoxCrossWebAppSPSiteSource, listBoxCrossWebAppSPSiteDest);
            SetLabelCrossWebAppPrompt();
        }

        private void comboBoxCrossWebAppManagedPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxCrossWebAppContentDBDest_SelectedIndexChanged(sender, e);
        }

        private void comboBoxCrossWebAppContentDBSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxCrossWebAppSource.SelectedItem == null)
                return;
            if (comboBoxCrossWebAppContentDBSource.SelectedItem == null)
                return;

            string strKeyCrossWebAppSource = ((KeyValuePair<string, string>)comboBoxCrossWebAppSource.SelectedItem).Key;
            string strKeyCrossWebAppContentDBSource = ((KeyValuePair<string, string>)comboBoxCrossWebAppContentDBSource.SelectedItem).Key;
            string strManagedPath = string.Empty;

            if (comboBoxCreateManagedPath.SelectedItem != null)
            {
                strManagedPath = ((KeyValuePair<string, string>)comboBoxCrossWebAppManagedPath.SelectedItem).Value;
            }
            PopulateSPSites(new Guid(strKeyCrossWebAppSource),
                new Guid(strKeyCrossWebAppContentDBSource),
                strManagedPath,
                listBoxCrossWebAppSPSiteSource);

            SPSiteAdmin2016.Properties.Settings.Default.CrossWebAppContentDBSource = strKeyCrossWebAppContentDBSource;
            SPSiteAdmin2016.Properties.Settings.Default.Save();

            ResetStatusListBoxSPSite(comboBoxCrossWebAppContentDBSource, comboBoxCrossWebAppContentDBDest, listBoxCrossWebAppSPSiteSource, listBoxCrossWebAppSPSiteDest);
            SetLabelCrossWebAppPrompt();
        }

        private void listBoxCrossWebAppSPSiteSource_DragDrop(object sender, DragEventArgs e)
        {
            listBoxCrossWebApp_DragDrop(sender, e, listBoxCrossWebAppSPSiteDest, listBoxCrossWebAppSPSiteSource);
        }

        private void listBoxCrossWebAppSPSiteSource_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void listBoxCrossWebAppSPSiteSource_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxCrossWebAppSPSiteSource.SelectedItem == null)
                return;
            listBoxCrossWebAppSPSiteSource.DoDragDrop(listBoxCrossWebAppSPSiteSource.SelectedItem, DragDropEffects.Move);
        }

        private void listBoxCrossWebAppSPSiteDest_DragDrop(object sender, DragEventArgs e)
        {
            listBoxCrossWebApp_DragDrop(sender, e, listBoxCrossWebAppSPSiteSource, listBoxCrossWebAppSPSiteDest);
        }

        private void listBoxCrossWebAppSPSiteDest_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void listBoxCrossWebAppSPSiteDest_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxCrossWebAppSPSiteDest.SelectedItem == null) return;
            listBoxCrossWebAppSPSiteDest.DoDragDrop(listBoxCrossWebAppSPSiteDest.SelectedItem, DragDropEffects.Move);
        }

        private void buttonCrossWebAppReset_Click(object sender, EventArgs e)
        {
            buttonClearPSScript_Click(null, null);
            comboBoxCrossWebAppContentDBSource_SelectedIndexChanged(null, null);
            comboBoxCrossWebAppContentDBDest_SelectedIndexChanged(null, null);
        }

        private void buttonCrossWebAppAll_Click(object sender, EventArgs e)
        {
            for (int i = listBoxCrossWebAppSPSiteSource.Items.Count - 1; i >= 0; i--)
            {
                listBoxCrossWebAppSPSiteSource.SelectedIndex = i;

                listBoxCrossWebApp_DragDrop(null, null, listBoxCrossWebAppSPSiteSource, listBoxCrossWebAppSPSiteDest);
            }
        }
    }
}
