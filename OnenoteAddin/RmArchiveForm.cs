using RemarkableSync.RmLine;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace RemarkableSync.OnenoteAddin
{
    public partial class RmArchiveForm : Form
    {
        class RmTreeNode : TreeNode
        {
            public RmTreeNode(string id, string visibleName, bool isCollection)
            {
                Text = (isCollection ? "\xD83D\xDCC1" : "\xD83D\xDCC4") + " " + visibleName;
                ID = id;
                VisibleName = visibleName;
                IsCollection = isCollection;
            }

            public string ID { get; set; }

            public string VisibleName { get; set; }

            public bool IsCollection { get; set; }

            public static List<RmTreeNode> FromRmItem(List<RmItem> rmItems)
            {
                List<RmTreeNode> nodes = new List<RmTreeNode>();
                foreach (var rmItem in rmItems)
                {
                    RmTreeNode node = new RmTreeNode(rmItem.ID, rmItem.VissibleName, rmItem.Type == RmItem.CollectionType);
                    node.Nodes.AddRange(FromRmItem(rmItem.Children).ToArray());
                    nodes.Add(node);
                }

                return nodes;
            }
        }

        private IRmDataSource _rmDataSource;
        private Application _application;
        private IConfigStore _configStore;
        private string _settingsRegPath;
        private List<RmTreeNode> rootItems;

        public RmArchiveForm(Application application, string settingsRegPath)
        {
            _settingsRegPath = settingsRegPath;
            _configStore = new WinRegistryConfigStore(_settingsRegPath);
            _application = application;

            InitializeComponent();
            InitializeData();            
        }

        private async void InitializeData()
        {
            rmTreeView.Nodes.Clear();
            lblInfo.Text = "Loading document list from reMarkable...";

            List<RmItem> rootItems = new List<RmItem>();

            try
            {
                await Task.Run(() =>
                {
                    int connMethod = -1;
                    try
                    {
                        string connMethodString = _configStore.GetConfig(SettingsForm.RmConnectionMethodConfig);
                        connMethod = Convert.ToInt32(connMethodString);
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine($"RmDownloadForm::RmDownloadForm() - Failed to get RmConnectionMethod config with err: {err.Message}");
                        // will default to cloud
                    }

                    switch (connMethod)
                    {
                        case (int)SettingsForm.RmConnectionMethod.Ssh:
                            _rmDataSource = new RmSftpDataSource(_configStore);
                            Console.WriteLine("Using SFTP data source");
                            break;
                        case (int)SettingsForm.RmConnectionMethod.RmCloud:
                        default:
                            _rmDataSource = new RmCloudDataSource(_configStore, new WinRegistryConfigStore(_settingsRegPath, false));
                            Console.WriteLine("Using rm cloud data source");
                            break;
                    }
                });
                rootItems = await _rmDataSource.GetItemHierarchy();
            }
            catch (Exception err)
            {
                Console.WriteLine($"Error getting notebook structure from reMarkable. Err: {err.Message}");
                MessageBox.Show($"Error getting notebook structure from reMarkable.\n{err.Message}", "Error");
                Close();
                return;
            }

            Console.WriteLine("Got item hierarchy from remarkable cloud");
            var treeNodeList = RmTreeNode.FromRmItem(rootItems);
            this.rootItems = treeNodeList;

            rmTreeView.Nodes.AddRange(treeNodeList.ToArray());
            Console.WriteLine("Added nodes to tree view");
            lblInfo.Text = "Select document to load into OneNote.";
            return;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private async void btnOk_Click(object sender, EventArgs e)
        {
            if (rmTreeView.SelectedNode == null && !chkArchiveAll.Checked)
            {
                MessageBox.Show(this, "No document selected.");
                return;
            }

            try {
                bool success = true;
                if (chkArchiveAll.Checked)
                {
                    foreach (RmTreeNode node in rootItems)
                    {
                        success &= await ImportDocuments(node);
                    }
                }
                else
                {
                    RmTreeNode rmTreeNode = (RmTreeNode)rmTreeView.SelectedNode;
                    Console.WriteLine($"Selected: {rmTreeNode.VisibleName} | {rmTreeNode.ID}");
                    success &= await ImportDocuments(rmTreeNode);
                }
                Console.WriteLine("Import " + (success ? "successful" : "failed"));

            }
            catch (Exception err) {
                Console.WriteLine($"Error importing document from reMarkable. Err: {err.Message}");
                MessageBox.Show($"Error importing document from reMarkable.\n{err.Message}", "Error");
                Close();
                return;
            }

            Close();
        }

        private async Task<bool> ImportDocuments(RmTreeNode rmTreeNode)
        {
            //MessageBox.Show("import multiple documents - " + rmTreeNode.VisibleName);
            bool result = true;

            if (rmTreeNode.IsCollection)
            {
                foreach (RmTreeNode n in rmTreeNode.Nodes)
                {

                    if (n.IsCollection)
                    {
                        result &= await ImportDocuments(n);
                    }
                    else
                    {
                        result &= await ImportDocument(n);
                    }
                }
            }
            else
            {
                result &= await ImportDocument(rmTreeNode);
            }

            return result;
        }

        private async Task<bool> ImportDocument(RmTreeNode rmTreeNode)
        {

            RmItem item = new RmItem();
            item.Type = RmItem.DocumentType;
            item.ID = rmTreeNode.ID;
            item.VissibleName = rmTreeNode.VisibleName;

            //MessageBox.Show("import a single document - " + item.VissibleName);

            List<RmPage> pages = new List<RmPage>();

            lblInfo.Text = $"Downloading {rmTreeNode.VisibleName}...";

            using (RmDownloadedDoc doc = await _rmDataSource.DownloadDocument(item))
            {
                Console.WriteLine("ImportDocument() - document downloaded");
                for (int i = 0; i < doc.PageCount; ++i)
                {
                    pages.Add(doc.GetPageContent(i));
                }
            }

            return await ImportContentAsTextAndImage(pages, rmTreeNode.VisibleName);
        }

        private async Task<bool> ImportContentAsTextAndImage(List<RmPage> pages, string visibleName)
        {
            List<string> textList = await ImportContentAsText(pages, visibleName);
            List<Bitmap> images = ImportContentAsGraphics(pages, visibleName);

            OneNoteHelper oneNoteHelper = new OneNoteHelper(_application);
            string currentSectionId = oneNoteHelper.GetCurrentSectionId();
            string newPageId = oneNoteHelper.CreatePage(currentSectionId, visibleName);

            lblInfo.Text = $"Appending text and images...";
            oneNoteHelper.AppendImagesAndText(newPageId, images, textList, 0.5);

            return true;
        }

        private async Task<List<String>> ImportContentAsText(List<RmPage> pages, string visibleName)
        {
            List<String> results = new List<String>();

            lblInfo.Text = $"Digitising {visibleName}...";
            MyScriptClient hwrClient = new MyScriptClient(_configStore);
            Console.WriteLine("ImportDocument() - requesting hand writing recognition");

            foreach (RmPage page in pages)
            {
                MyScriptResult result = await hwrClient.RequestHwr(new List<RmPage> {page});
                results.Add(result.label);
            }

            return results;
        }

        private void UpdateOneNoteWithHwrResult(string name, MyScriptResult result)
        {
            OneNoteHelper oneNoteHelper = new OneNoteHelper(_application);
            string currentSectionId = oneNoteHelper.GetCurrentSectionId();
            string newPageId = oneNoteHelper.CreatePage(currentSectionId, name);
            oneNoteHelper.AddPageContent(newPageId, result.label);
        }

        private List<Bitmap> ImportContentAsGraphics(List<RmPage> pages, string visibleName)
        {
            lblInfo.Text = $"Importing {visibleName} as graphics...";
            OneNoteHelper oneNoteHelper = new OneNoteHelper(_application);
            string currentSectionId = oneNoteHelper.GetCurrentSectionId();

            List<Bitmap> images = RmLinesDrawer.DrawPages(pages);

            return images;
        }

        private void RmDownloadForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _rmDataSource?.Dispose();
        }

        private void rmTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void chkArchiveAll_CheckedChanged(object sender, EventArgs e)
        {
            if (chkArchiveAll.Checked)
            {
                rmTreeView.Enabled = false;
            } else
            {
                rmTreeView.Enabled = true;
            }
            
        }
    }
}
