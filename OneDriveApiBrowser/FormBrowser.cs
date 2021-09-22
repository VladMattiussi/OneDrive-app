// ------------------------------------------------------------------------------
//  Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

namespace OneDriveApiBrowser
{
    using Aspose.Words;
    using Aspose.Pdf;
    using Aspose.Cells;
    using Microsoft.Graph;
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.IO;
    using System.Collections;
    using System.Linq;
    using Aspose.Slides;
    using System.Threading;

    public partial class FormBrowser : Form
    {
        public const string MsaClientId = "6d3f25d0-553f-4a5b-912e-434b790bc6d8";
        public const string MsaReturnUrl = "urn:ietf:wg:oauth:2.0:oob";


        private enum ClientType
        {
            Consumer,
            Business
        }

        private const int UploadChunkSize = 10 * 1024 * 1024;       // 10 MB
        //private IOneDriveClient oneDriveClient { get; set; }
        private GraphServiceClient graphClient { get; set; }
        private ClientType clientType { get; set; }
        private DriveItem CurrentFolder { get; set; }
        private DriveItem SelectedItem { get; set; }

        private OneDriveTile _selectedTile;

        public FormBrowser()
        {
            InitializeComponent();
        }

        private void ShowWork(bool working)
        {
            this.UseWaitCursor = working;
            this.progressBar1.Visible = working;

        }

        private async Task LoadFolderFromId(string id)
        {
            if (null == this.graphClient) return;

            // Update the UI for loading something new
            ShowWork(true);
            LoadChildren(new DriveItem[0]);

            try
            {
                var expandString = this.clientType == ClientType.Consumer
                    ? "thumbnails,children($expand=thumbnails)"
                    : "thumbnails,children";

                var folder =
                    await this.graphClient.Drive.Items[id].Request().Expand(expandString).GetAsync();

                ProcessFolder(folder);
            }
            catch (Exception exception)
            {
                PresentServiceException(exception);
            }

            ShowWork(false);
        }

        private async Task LoadFolderFromPath(string path = null)
        {
            if (null == this.graphClient) return;

            // Update the UI for loading something new
            ShowWork(true);
            LoadChildren(new DriveItem[0]);

            try
            {
                DriveItem folder;

                var expandValue = this.clientType == ClientType.Consumer
                    ? "thumbnails,children($expand=thumbnails)"
                    : "thumbnails,children";

                if (path == null)
                {
                    folder = await this.graphClient.Drive.Root.Request().Expand(expandValue).GetAsync();
                }
                else
                {
                    folder =
                        await
                            this.graphClient.Drive.Root.ItemWithPath("/" + path)
                                .Request()
                                .Expand(expandValue)
                                .GetAsync();
                }

                ProcessFolder(folder);
            }
            catch (Exception exception)
            {
                PresentServiceException(exception);
            }

            ShowWork(false);
        }

        private void ProcessFolder(DriveItem folder)
        {
            if (folder != null)
            {
                this.CurrentFolder = folder;

                LoadProperties(folder);

                if (folder.Folder != null && folder.Children != null && folder.Children.CurrentPage != null)
                {
                    LoadChildren(folder.Children.CurrentPage);
                }
            }
        }

        private void LoadProperties(DriveItem item)
        {
            this.SelectedItem = item;
            objectBrowser.SelectedItem = item;
        }

        private void LoadChildren(IList<DriveItem> items)
        {
            flowLayoutContents.SuspendLayout();
            flowLayoutContents.Controls.Clear();

            // Load the children
            foreach (var obj in items)
            {
                AddItemToFolderContents(obj);
            }

            flowLayoutContents.ResumeLayout();
        }

        private void AddItemToFolderContents(DriveItem obj)
        {
            flowLayoutContents.Controls.Add(CreateControlForChildObject(obj));
        }

        private void RemoveItemFromFolderContents(DriveItem itemToDelete)
        {
            flowLayoutContents.Controls.RemoveByKey(itemToDelete.Id);
        }

        private System.Windows.Forms.Control CreateControlForChildObject(DriveItem item)
        {
            int n = 0;
            if (item.Name.Contains("."))
            {

                if (item.Name.Substring(Math.Max(0, item.Name.Length - 4)).Contains("doc"))
                {
                    n = 1;
                }
                if (item.Name.Substring(Math.Max(0, item.Name.Length - 4)).Contains("pdf"))
                {
                    n = 2;
                }
                if (item.Name.Substring(Math.Max(0, item.Name.Length - 4)).Contains("xlsx"))
                {
                    n = 3;
                }
                if (item.Name.Substring(Math.Max(0, item.Name.Length - 4)).Contains("pptx"))
                {
                    n = 4;
                }
            }
            OneDriveTile tile = new OneDriveTile(this.graphClient,n);
            tile.SourceItem = item;
            tile.Click += ChildObject_Click;
            tile.DoubleClick += ChildObject_DoubleClick;
            tile.Name = item.Id;
            return tile;
        }

        void ChildObject_DoubleClick(object sender, EventArgs e)
        {
            var item = ((OneDriveTile)sender).SourceItem;

            // Look up the object by ID
            NavigateToFolder(item);
        }
        void ChildObject_Click(object sender, EventArgs e)
        {
            if (null != _selectedTile)
            {
                _selectedTile.Selected = false;
            }
            
            var item = ((OneDriveTile)sender).SourceItem;
            LoadProperties(item);
            _selectedTile = (OneDriveTile)sender;
            _selectedTile.Selected = true;
        }

        private void FormBrowser_Load(object sender, EventArgs e)
        {
            
        }

        private void NavigateToFolder(DriveItem folder)
        {
            Task t = LoadFolderFromId(folder.Id);

            // Fix up the breadcrumbs
            var breadcrumbs = flowLayoutPanelBreadcrumb.Controls;
            bool existingCrumb = false;
            foreach (LinkLabel crumb in breadcrumbs)
            {
                if (crumb.Tag == folder)
                {
                    RemoveDeeperBreadcrumbs(crumb);
                    existingCrumb = true;
                    break;
                }
            }

            if (!existingCrumb)
            {
                LinkLabel label = new LinkLabel();
                label.Text = "> " + folder.Name;
                label.LinkArea = new LinkArea(2, folder.Name.Length);
                label.LinkClicked += linkLabelBreadcrumb_LinkClicked;
                label.AutoSize = true;
                label.Tag = folder;
                flowLayoutPanelBreadcrumb.Controls.Add(label);
            }
        }

        private void linkLabelBreadcrumb_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel link = (LinkLabel)sender;

            RemoveDeeperBreadcrumbs(link);

            DriveItem item = link.Tag as DriveItem;
            if (null == item)
            {

                Task t = LoadFolderFromPath(null);
            }
            else
            {
                Task t = LoadFolderFromId(item.Id);
            }
        }

        private void RemoveDeeperBreadcrumbs(LinkLabel link)
        {
            // Remove the breadcrumbs deeper than this item
            var breadcrumbs = flowLayoutPanelBreadcrumb.Controls;
            int indexOfControl = breadcrumbs.IndexOf(link);
            for (int i = breadcrumbs.Count - 1; i > indexOfControl; i--)
            {
                breadcrumbs.RemoveAt(i);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UpdateConnectedStateUx(bool connected)
        {
            signInMsaToolStripMenuItem.Visible = !connected;
            signOutToolStripMenuItem.Visible = connected;
            flowLayoutPanelBreadcrumb.Visible = connected;
            flowLayoutContents.Visible = connected;
        }

        private async void signInMsaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await this.SignIn();
            //Action a = () => SearchUpload();
            //Task task = Task.Run(a);
            Thread thread = new Thread(() => SearchUploadAsync());
            thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
            thread.Start();
        }

        [STAThread]
        private async Task SearchUploadAsync()
        {
            string OriginPath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            OriginPath = OriginPath.Remove(OriginPath.Length - 22, 22);
            OriginPath = OriginPath + "data.txt";
            OriginPath = new Uri(OriginPath).LocalPath;
            if (!System.IO.File.Exists(OriginPath))
            {
                using (StreamWriter sw = System.IO.File.CreateText(OriginPath))
                {

                }
            }

            string folder;
            string categoria = null;
            ArrayList lista = new ArrayList();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            var response = dialog.ShowDialog();
            DirectoryInfo d = new DirectoryInfo(dialog.SelectedPath);
            DirectoryInfo sub = d.CreateSubdirectory("File_Caricati");
            while (true)
            {
                FileInfo[] list = d.GetFiles();
                foreach (var file in list)
                {
                    lista = FolderChooser(file.FullName, OriginPath);
                    if (lista != null)
                    {
                        ArrayList tags = new ArrayList();
                        categoria = (string)lista[1];
                        if (categoria == null)
                        {
                            categoria = "";
                        }
                        tags = (ArrayList)lista[0];
                        foreach (string f in tags)
                        {
                            folder = f;
                            if (folder != null)
                                if (folder != null)
                                {
                                    using (var stream = new System.IO.FileStream(file.FullName, System.IO.FileMode.Open))
                                    {
                                        if (stream != null)
                                        {
                                            // Since the ItemWithPath method is available only at Drive.Root, we need to strip
                                            // /drive/root: (12 characters) from the parent path string.
                                            // string folderPath = targetFolder.ParentReference == null
                                            // ? ""
                                            // : targetFolder.ParentReference.Path.Remove(0, 12) + "/" + Uri.EscapeUriString(targetFolder.Name);
                                            var uploadPath = folder + "/" + categoria + "/" + Uri.EscapeUriString(System.IO.Path.GetFileName(file.Name));

                                            try
                                            {
                                                var uploadedItem =
                                                    await
                                                        this.graphClient.Drive.Root.ItemWithPath(uploadPath).Content.Request().PutAsync<DriveItem>(stream);

                                                //AddItemToFolderContents(uploadedItem);
                                            }
                                            catch (Exception exception)
                                            {
                                                PresentServiceException(exception);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("File non caricato,scegliere un file doc,pdf,xlsx o pptx");
                                }
                        }
                        string brawl = Path.Combine(sub.FullName, file.Name);
                        file.CopyTo(brawl, true);
                        file.Delete();
                    }
                }
            }
        }
        
        private async Task SignIn()
        {

            try
            {
                this.graphClient = AuthenticationHelper.GetAuthenticatedClient();
            }
            catch (ServiceException exception)
            {

             PresentServiceException(exception);

            }

            try
            {
                await LoadFolderFromPath();

                UpdateConnectedStateUx(true);
            }
            catch (ServiceException exception)
            {
                PresentServiceException(exception);
                this.graphClient = null;
            }
        }

        private void signOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.graphClient != null)
            {
                AuthenticationHelper.SignOut();
            }

            UpdateConnectedStateUx(false);
        }

        private String GetFileStreamForUpload(string targetFolderName, out string originalFilename,out string FilePath)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Upload to " + targetFolderName;
            dialog.Filter = "All Files (*.*)|*.*";
            dialog.CheckFileExists = true;
            var response = dialog.ShowDialog();
            if (response != DialogResult.OK)
            {
                originalFilename = null;
                FilePath = null;
                return null;
            }

            try
            {
                originalFilename = System.IO.Path.GetFileName(dialog.FileName);
                FilePath = dialog.FileName;

                return dialog.FileName;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error uploading file: " + ex.Message);
                originalFilename = null;
                FilePath = null;
                return null;
            }
        }

        private async void simpleUploadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string OriginPath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase; 
            OriginPath = OriginPath.Remove(OriginPath.Length - 22, 22);
            OriginPath = OriginPath + "data.txt";
            OriginPath = new Uri(OriginPath).LocalPath;
            if (!System.IO.File.Exists(OriginPath))
            {
            
            using (StreamWriter sw = System.IO.File.CreateText(OriginPath))
                {

                }
            }
            DialogResult dialogResult = MessageBox.Show("Vuoi caricare più di 1 file?", "Caricamento multiplo", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string folder;
                string categoria = null;
                ArrayList lista = new ArrayList();
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                var response = dialog.ShowDialog();
                if (dialog.SelectedPath != null && dialog.SelectedPath != "")
                {
                    DirectoryInfo d = new DirectoryInfo(dialog.SelectedPath);
                    FileInfo[] list = d.GetFiles();
                    foreach (var file in list)
                    {
                        lista = FolderChooser(file.FullName, OriginPath);
                        if (lista != null)
                        {
                            ArrayList tags = new ArrayList();
                            categoria = (string)lista[1];
                            if (categoria == null)
                            {
                                categoria = "";
                            }
                            tags = (ArrayList)lista[0];
                            foreach (string f in tags)
                            {
                                folder = f;
                                if (folder != null)
                                    if (folder != null)
                                    {
                                        using (var stream = new System.IO.FileStream(file.FullName, System.IO.FileMode.Open))
                                        {
                                            if (stream != null)
                                            {
                                                // Since the ItemWithPath method is available only at Drive.Root, we need to strip
                                                // /drive/root: (12 characters) from the parent path string.
                                                // string folderPath = targetFolder.ParentReference == null
                                                // ? ""
                                                // : targetFolder.ParentReference.Path.Remove(0, 12) + "/" + Uri.EscapeUriString(targetFolder.Name);
                                                var uploadPath = folder + "/" + categoria + "/" + Uri.EscapeUriString(System.IO.Path.GetFileName(file.Name));

                                                try
                                                {
                                                    var uploadedItem =
                                                        await
                                                            this.graphClient.Drive.Root.ItemWithPath(uploadPath).Content.Request().PutAsync<DriveItem>(stream);

                                                    //AddItemToFolderContents(uploadedItem);

                                                    MessageBox.Show("Caricato in " + folder);
                                                }
                                                catch (Exception exception)
                                                {
                                                    PresentServiceException(exception);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("File non caricato,scegliere un file doc,pdf,xlsx o pptx");
                                    }
                            }
                        }
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                var targetFolder = this.CurrentFolder;
                string folder = null;
                string FilePath = "";
                string filename;
                string categoria = null;
                ArrayList lista = new ArrayList();

                GetFileStreamForUpload(folder, out filename, out FilePath);
                lista = FolderChooser(FilePath, OriginPath);
                if (lista != null)
                {
                    ArrayList tags = new ArrayList();
                    categoria = (string)lista[1];
                    if(categoria == null)
                    {
                        categoria = "";
                    }
                    tags = (ArrayList)lista[0];
                    foreach (string f in tags)
                    {
                        folder = f;
                        if (folder != null)
                        {
                            using (var stream = new System.IO.FileStream(FilePath, System.IO.FileMode.Open))
                            {
                                if (stream != null)
                                {
                                    // Since the ItemWithPath method is available only at Drive.Root, we need to strip
                                    // /drive/root: (12 characters) from the parent path string.
                                    // string folderPath = targetFolder.ParentReference == null
                                    // ? ""
                                    // : targetFolder.ParentReference.Path.Remove(0, 12) + "/" + Uri.EscapeUriString(targetFolder.Name);
                                    var uploadPath = folder + "/" + categoria + "/" + Uri.EscapeUriString(System.IO.Path.GetFileName(filename));

                                    try
                                    {
                                        var uploadedItem =
                                            await
                                                this.graphClient.Drive.Root.ItemWithPath(uploadPath).Content.Request().PutAsync<DriveItem>(stream);

                                        AddItemToFolderContents(uploadedItem);

                                        //MessageBox.Show("Uploaded with ID: " + uploadedItem.Id);
                                        MessageBox.Show("Caricato in " + folder);
                                    }
                                    catch (Exception exception)
                                    {
                                        PresentServiceException(exception);
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("File non caricato,scegliere un file doc,pdf,xlsx o pptx");
                        }
                    }
                }
            }       
        }

        private ArrayList FolderChooser(string filepath,string origin)
        {
            if (filepath != null)
            {
                if (filepath.Substring(Math.Max(0, filepath.Length - 4)).Contains("doc"))
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(filepath);
                    if (doc.BuiltInDocumentProperties.Contains("Keywords"))
                    {
                        string keys = doc.BuiltInDocumentProperties.Keywords;
                        return CheckTag(keys, origin);
                    }
                }
                if (filepath.Substring(Math.Max(0, filepath.Length - 4)).Contains("pdf"))
                {
                    Aspose.Pdf.Document pdf = new Aspose.Pdf.Document(filepath);
                    if (pdf.Info.ContainsKey("Keywords"))
                    {
                        string keys = pdf.Info["Keywords"];
                        return CheckTag(keys, origin);
                    }
                }
                if (filepath.Substring(Math.Max(0, filepath.Length - 4)).Contains("xlsx"))
                {
                    Aspose.Cells.Workbook cell = new Aspose.Cells.Workbook(filepath);
                    if (cell.BuiltInDocumentProperties.Contains("Keywords"))
                    {
                        string keys = cell.BuiltInDocumentProperties.Keywords;
                        return CheckTag(keys, origin);
                    }
                }
                if (filepath.Substring(Math.Max(0, filepath.Length - 4)).Contains("pptx"))
                {
                    Presentation pres = new Presentation(filepath);
                    if (pres.DocumentProperties.Keywords != null)
                    {
                        string keys = pres.DocumentProperties.Keywords;
                        return CheckTag(keys, origin);
                    }
                }
            }
            return null;
        }

        private String ErrorMessage(string key,string origin)
        {
            DialogResult dialogResult = MessageBox.Show(" Tag "+ key + " non trovato, vuoi comunque caricare il file in una nuova cartella?", "Tag non trovato", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                using (StreamWriter sw = System.IO.File.AppendText(origin))
                {
                    sw.WriteLine(key);
                }
                return key;
            }
            else
            {
                return null;
            }
        }

        private ArrayList CheckTag(String tag,String origin)
        {
            ArrayList a0 = new ArrayList();
            ArrayList a1 = new ArrayList();
            int count = tag.Count(x => x == ';');
            int n;
            string sub = null;
            for (int a = 0; a <= count; a++)
            {
                n = 0;
                string key = tag.Split(';').First();
                tag = tag.Substring(tag.IndexOf(';') + 1);
                if (key == ("SD"))
                {
                    a1.Add("Sistemidata");
                    n = 1;
                }
                if (key == ("SP"))
                {
                    a1.Add("Service Point");
                    n = 1;
                }
                if (key == ("MP"))
                {
                    a1.Add("Master Point");
                    n = 1;
                }
                if (key == ("PP"))
                {
                    a1.Add("Partner Program");
                    n = 1;
                }
                if (key == ("CD"))
                {
                    a1.Add("Commerciali Diretta");
                    n = 1;
                }

                if (key.Contains("IC"))
                {
                    sub = ("Informativa Commerciale");
                    n = 1;
                }
                if (key.Contains("LP"))
                {
                    sub = ("Listino Prezzi");
                    n = 1;
                }
                if (key.Contains("FC"))
                {
                    sub = ("Format Contratti");
                    n = 1;
                }
                if (key.Contains("SLA"))
                {
                    sub = ("Service Level Agreement");
                    n = 1;
                }
                if (key.Contains("IP"))
                {
                    sub = ("Istruzioni Operative");
                    n = 1;
                }
                if (key.Contains("KIT"))
                {
                    sub = ("Kit di rivendita");
                    n = 1;
                }
                if (key.Contains("RA"))
                {
                    sub = ("Modulo RA");
                    n = 1;
                }
                if (key.Contains("A1"))
                {
                    sub = ("Modulo A1");
                    n = 1;
                }
                if (key.Contains("SCP"))
                {
                    sub = ("Scheda Prodotto");
                    n = 1;
                }
                if (key.Contains("SCT"))
                {
                    sub = ("Scheda Tecnica");
                    n = 1;
                }
                if (key.Contains("REP"))
                {
                    sub = ("Report Personalizzato");
                    n = 1;
                }
                if (key.Contains("LEAD"))
                {
                    sub = ("Lead Assegnati");
                    n = 1;
                }
                if (key.Contains("CAN"))
                {
                    sub = ("Canvas Personalizzati");
                    n = 1;
                }
                if (key.Contains("MKTG"))
                {
                    sub = ("Materiale Marketing Vario");
                    n = 1;
                }
                if (key.Contains("PRY"))
                {
                    sub = ("Documentazione Privacy Varia");
                    n = 1;
                }
                if (key.Contains("COM"))
                {
                    sub = ("Documentazione Commerciale Varia");
                    n = 1;
                }
                if (key.Contains("AMM"))
                {
                    sub = ("Documentazione Amministrativa Varia");
                    n = 1;
                }

                using (StreamReader sr = System.IO.File.OpenText(origin))
                {
                    string s = "";
                    while ((s = sr.ReadLine()) != null)
                    {
                        if (key == s)
                        {
                            a1.Add(s);
                            n = 1;
                        }
                    }
                }
                if (n == 0)
                {
                    a1.Add(ErrorMessage(key, origin));
                    n = 1;
                }
            }
            a0.Add(a1);
            a0.Add(sub);
            return a0;
    }

        private async void simpleIDbasedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var targetFolder = this.CurrentFolder;
            string FilePath;
            string folder = "";

            string filename;
            using (var stream = new System.IO.FileStream(GetFileStreamForUpload(folder, out filename, out FilePath), System.IO.FileMode.Open))
            {
                if (stream != null)
                {
                    try
                    {
                        var uploadedItem =
                            await
                                this.graphClient.Drive.Items[targetFolder.Id].ItemWithPath(filename).Content.Request()
                                    .PutAsync<DriveItem>(stream);

                        AddItemToFolderContents(uploadedItem);

                        MessageBox.Show("Uploaded with ID: " + uploadedItem.Id);
                    }
                    catch (Exception exception)
                    {
                        PresentServiceException(exception);
                    }
                }
            }
        }

        private async void createFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormInputDialog dialog = new FormInputDialog("Create Folder", "New folder name:");
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(dialog.InputText))
            {
                try
                {
                    var folderToCreate = new DriveItem { Name = dialog.InputText, Folder = new Folder() };
                    var newFolder =
                        await this.graphClient.Drive.Items[this.SelectedItem.Id].Children.Request()
                            .AddAsync(folderToCreate);

                    if (newFolder != null)
                    {
                        MessageBox.Show("Created new folder with ID " + newFolder.Id);
                        this.AddItemToFolderContents(newFolder);
                    }
                }
                catch(ServiceException exception)
                {
                    PresentServiceException(exception);

                }
                catch (Exception exception)
                {
                    PresentServiceException(exception);
                }
            }
        }

        private static void PresentServiceException(Exception exception)
        {
            string message = null;
            var oneDriveException = exception as ServiceException;
            if (oneDriveException == null)
            {
                message = exception.Message;
            }
            else
            {
                message = string.Format("{0}{1}", Environment.NewLine, oneDriveException.ToString());
            }

            MessageBox.Show(string.Format("OneDrive reported the following error: {0}", message));
        }

        private async void deleteSelectedItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var itemToDelete = this.SelectedItem;
            var result = MessageBox.Show("Sei sicuro di voler eliminare " + itemToDelete.Name + "?", "Conferma eliminazione", MessageBoxButtons.YesNo);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    await this.graphClient.Drive.Items[itemToDelete.Id].Request().DeleteAsync();
                    
                    RemoveItemFromFolderContents(itemToDelete);
                    MessageBox.Show("Oggetto eliminato con successo");
                }
                catch (Exception exception)
                {
                    PresentServiceException(exception);
                }
            }
        }

        private async void getChangesHereToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var result =
                    await this.graphClient.Drive.Items[this.CurrentFolder.Id].Delta().Request().GetAsync();

                foreach ( DriveItem item in result)
                {
                    Console.WriteLine(item.Name);
                }
            }
            catch (Exception ex)
            {
                PresentServiceException(ex);
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private async void saveSelectedFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var item = this.SelectedItem;
            if (null == item)
            {
                MessageBox.Show("Nothing selected.");
                return;
            }

            var dialog = new SaveFileDialog();
            dialog.FileName = item.Name;
            dialog.Filter = "All Files (*.*)|*.*";
            var result = dialog.ShowDialog();
            if (result != System.Windows.Forms.DialogResult.OK)
                return;

            using (var stream = await this.graphClient.Drive.Items[item.Id].Content.Request().GetAsync())
            using (var outputStream = new System.IO.FileStream(dialog.FileName, System.IO.FileMode.Create))
            {
                await stream.CopyToAsync(outputStream);
            }
        }
    }
}
