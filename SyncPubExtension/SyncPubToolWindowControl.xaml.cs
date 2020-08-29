namespace SyncPubExtension
{
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Xml;

    /// <summary>
    /// Interaction logic for SyncPubToolWindowControl.
    /// </summary>
    public partial class SyncPubToolWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncPubToolWindowControl"/> class.
        /// </summary>
        public SyncPubToolWindowControl()
        {
            this.InitializeComponent();
        }
        Project project = null;
        OleDbConnection connection = null;
        OleDbCommand command = null;
        bool usePublish;
        string _p, rootPath, profileName;
        DTE _applicationObject;
        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        public void PublishProjectInZip(object sender, RoutedEventArgs e)
        {
            logobj.Text = string.Empty;
            project = null;
            if (project is null)
                project = GetActiveProject();

            if (project is null)
            {
                llog("please open project !");
                return;
            }

            _applicationObject = project.DTE;

            if (!CreateNewAccessDatabaseOrOpen($@"{project.FileName}"))
            {
                llog("Error !");
                return;
            }



            llog("Find project type...");

            _p = new FileInfo(project.FileName).Directory.FullName;

            string xml = File.ReadAllText(project.FileName);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            string json = JsonConvert.SerializeXmlNode(doc);

            dynamic obj = JsonConvert.DeserializeObject(json);
            var type = obj.Project.PropertyGroup[0].OutputType.Value;
            llog($@"Current Project Type {type}");
            rootPath = string.Empty;
            usePublish = false;
            switch (type)
            {
                case "WinExe":
                    llog($@"This Project be Winform/WPF project");
                    if (obj.Project.PropertyGroup[0].PublishWizardCompleted != null)
                    {
                        var isCompleted = obj.Project.PropertyGroup[0].PublishWizardCompleted.Value;
                        var publishUrl = obj.Project.PropertyGroup[0].PublishUrl.Value;
                        llog($@"PublishWizardCompleted : {isCompleted}");
                        var rs = System.Windows.Forms.MessageBox.Show($@"Yes : use publish folder{Environment.NewLine}No : use {project.ConfigurationManager.ActiveConfiguration.ConfigurationName} folder name", "Select Folder", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Asterisk);
                        if (rs == System.Windows.Forms.DialogResult.Yes)
                        {
                            rootPath = publishUrl;
                            usePublish = true;
                        }
                        else
                            rootPath = $@"{_p}\bin\{project.ConfigurationManager.ActiveConfiguration.ConfigurationName}";
                    }
                    else
                    {
                        llog("project didn't publish so use debug folder for publish folder");
                        rootPath = $@"{_p}\bin\{project.ConfigurationManager.ActiveConfiguration.ConfigurationName}";
                    }
                    break;
                case "Exe":
                    llog($@"This Project be Console project");
                    rootPath = $@"{_p}\bin\{project.ConfigurationManager.ActiveConfiguration.ConfigurationName}";
                    break;
                case "Library":
                    llog($@"This Project be Dll (including web applications) project");
                    if (!File.Exists(project.FileName + ".user"))
                    {
                        llog("Please! publish to project");
                        return;
                    }
                    string xmlUser = File.ReadAllText(project.FileName + ".user");
                    XmlDocument docUser = new XmlDocument();
                    docUser.LoadXml(xmlUser);

                    string jsonUser = JsonConvert.SerializeXmlNode(docUser);

                    dynamic objUser = JsonConvert.DeserializeObject(jsonUser);

                    if (objUser.Project.PropertyGroup.NameOfLastUsedPublishProfile is null)
                    {
                        llog($@"Please! Create folder profile and publish min 1 times");
                        return;
                    }
                    else
                    {
                        profileName = objUser.Project.PropertyGroup.NameOfLastUsedPublishProfile;
                        var profilePath = Directory.GetFiles(_p, $@"{profileName}.pubxml.user", SearchOption.AllDirectories)?.FirstOrDefault();
                        if (profilePath is null || profilePath == string.Empty)
                        {
                            llog($@"Profile not found !!");
                            return;
                        }
                        else
                        {
                            string xmlProfile = File.ReadAllText(profilePath);
                            XmlDocument docProfile = new XmlDocument();
                            docProfile.LoadXml(xmlProfile);

                            string jsonProfile = JsonConvert.SerializeXmlNode(docProfile);

                            dynamic objProfile = JsonConvert.DeserializeObject(jsonProfile);

                            if (objProfile.Project.PropertyGroup._PublishTargetUrl is null)
                            {
                                llog($@"Profile found but not published !!");
                                return;
                            }
                            else
                            {
                                rootPath = objProfile.Project.PropertyGroup._PublishTargetUrl.Value;
                                usePublish = true;
                            }
                        }
                    }
                    break;
                default:
                    llog($@"[{type}] Unsupported type. returning...");
                    return;
            }
            if (rootPath is "")
            {
                llog("Error !");
                return;
            }
            var temp = "";


            var sb2 = project.DTE.Solution.SolutionBuild as SolutionBuild2;
            //ChangeProjectContexts(project, "Release");
            //sb2.BuildProject("Release", project.UniqueName, true);
            //sb2.PublishProject("Release", project.UniqueName, true);
            //return;
            //System.Diagnostics.Process.Start("msbuild.exe", $@"""{project.Name}.csproj"" /p:PublishProfile=""{profileName}.pubxml"" /p:DeployOnBuild=true /p:VisualStudioVersion=""14.0""");

            //if (usePublish)
            //{
            //    //project.DTE.ExecuteCommand("ClassViewContextMenus.ClassViewProject.Publish");
            //    //llog("Publishing");
            //    ////sb2.Publish(true);
            //    ////sb2.Run();
            //    ////sb2.BuildProject(project.ConfigurationManager.ActiveConfiguration.ConfigurationName, project.UniqueName, true);
            //    //sb2.Publish(true);
            //    ////sb2.PublishProject(project.ConfigurationManager.ActiveConfiguration.ConfigurationName, project.UniqueName, true);
            //    //llog("Published");
            //    ////System.Diagnostics.Process.Start("msbuild", $@"""{project.UniqueName.Split('\\')[1]}"" /p:PublishProfile=""{profileName}.pubxml"" /p:DeployOnBuild=true /p:VisualStudioVersion=""14.0""");
            //}
            if (!usePublish)
            {
                llog("Building");
                sb2.Build(true);
                llog("Builded");
            }

            llog("Founding and ziping...");
            connection.Open();
            Founding();
            connection.Close();
            llog("Finish");

        }
        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        public void OpenPublishWindow(object sender, RoutedEventArgs e)
        {
            project = GetActiveProject();
            if (project is null)
            {
                llog("Please select any project");
                return;
            }

            project.DTE.ExecuteCommand("ClassViewContextMenus.ClassViewProject.Publish");
        }
        private void ChangeProjectContexts(EnvDTE.Project project, string configurationName)
        {
            EnvDTE.SolutionConfigurations solutionConfigurations;

            solutionConfigurations = _applicationObject.Solution.SolutionBuild.SolutionConfigurations;

            foreach (EnvDTE80.SolutionConfiguration2 solutionConfiguration2 in solutionConfigurations)
            {
                foreach (EnvDTE.SolutionContext solutionContext in solutionConfiguration2.SolutionContexts)
                {
                    if (solutionContext.ProjectName == project.UniqueName)
                    {
                        solutionContext.ConfigurationName = configurationName;
                    }
                }
            }
        }
        #region Support Method
        public class Files
        {
            public string Id { get; set; }
            public string Path { get; set; }
            public string LastWriteTime { get; set; }
        }
        private List<string> ChangeFilePaths { get; set; }
        private DataTable FileLogs { get; set; }
        private void FileLogsRestart()
        {
            string query = $@"Select * From {nameof(Files)}";
            command.CommandText = query;
            var asd = connection.GetSchema();
            var reader = command.ExecuteReader();
            reader.Read();
            DataTable dataTable = new DataTable();
            dataTable.Load(reader);
            FileLogs = dataTable;
        }
        private void Founding()
        {
            FileLogsRestart();
            ChangeFilePaths = new List<string>();
            var allFilesInFromPath = Directory.GetFiles(rootPath, "*", SearchOption.AllDirectories);
            foreach (var fromFilePath in allFilesInFromPath)
            {
                var fromFileInfo = new FileInfo(fromFilePath);

                //var fileLog = FileLogs?.FirstOrDefault(fl => fl.Path == fromFilePath);
                var fileLog = (from log in FileLogs.Rows.Cast<DataRow>().ToList()
                               where log["Path"].ToString() == fromFilePath
                               select log).FirstOrDefault();

                bool IsCorrect = false;
                if (fileLog != null)
                    IsCorrect = fileLog["LastWriteTime"].ToString() != fromFileInfo.LastWriteTime.ToString("yyyyMMddHHmmss");

                string queryForUpdate = string.Empty;
                if (fileLog != null)
                    queryForUpdate = $@"
                    Update
                        {nameof(Files)}
                    Set
                        LastWriteTime = '{fromFileInfo.LastWriteTime.ToString("yyyyMMddHHmmss")}'
                    Where
                        Id = '{fileLog["Id"]}'
                    ";
                string queryForInsert = string.Empty;
                if (fileLog is null)
                    queryForInsert = $@"
                    Insert Into {nameof(Files)}
                    (
                    Id
                    ,Path
                    ,LastWriteTime
                    )
                    Values
                    (
                    '{Guid.NewGuid().ToString()}'
                    ,'{fromFilePath}'
                    ,'{fromFileInfo.LastWriteTime.ToString("yyyyMMddHHmmss")}'
                    )
                    ";

                if (IsCorrect)
                {
                    AddList();
                    command.CommandText = queryForUpdate;
                    command.ExecuteNonQuery();
                }
                else if (fileLog is null)
                {
                    AddList();
                    command.CommandText = queryForInsert;
                    command.ExecuteNonQuery();
                }

                void AddList()
                {
                    ChangeFilePaths.Add(fromFilePath);
                }
            }
            if (ChangeFilePaths.Count == 0)
                llog("Change file(s) not found");
            else
                CreateZip();
        }
        private void CreateZip()
        {
            FileLogsRestart();
            string zipPath = $@"{_p}\publishs\";
            if (!Directory.Exists(zipPath))
                Directory.CreateDirectory(zipPath);
            string createZipPath = $@"{zipPath}\{project.Name}-{DateTime.Now.ToString("yyyy.MM.dd--HH.mm.ss")}.zip";

            FileStream createZipStream = new FileStream(createZipPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            ZipArchive createZip = new ZipArchive(createZipStream, ZipArchiveMode.Update, true, Encoding.UTF8);

            foreach (var filePath in ChangeFilePaths)
            {
                createZip.CreateEntryFromFile(filePath, $@"{GetRootPath(filePath, rootPath).Remove(0, 1)}", CompressionLevel.Optimal);
            }

            createZip.Dispose();
            createZipStream.Close();

            llog($@"Create Zip. Name : {createZipPath}");

            System.Diagnostics.Process.Start(zipPath);
        }
        private string GetRootPath(string path, string root) => path.Replace(root, "");
        public bool CreateNewAccessDatabaseOrOpen(string fileName)
        {
            fileName += ".accdb";
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5";
            if (!File.Exists(fileName))
            {
                if (Create())
                {
                    llog($@"Access Database created. Path : {fileName}");
                }
                else
                {
                    llog("Access Database file not created.");
                    return false;
                }
            }
            connection = new OleDbConnection(connectionString);
            command = new OleDbCommand("", connection);
            llog("create connection");
            return true;
            bool Create()
            {
                bool result = false;

                ADOX.Catalog cat = new ADOX.Catalog();
                //ADOX.Table table = new ADOX.Table
                //{
                //    //Create the table and it's fields. 
                //    Name = nameof(Files)
                //};
                ///*
                // Buraya bakıcaksın unutma tablo yapısına
                // */
                //table.Columns.Append("Id", ADOX.DataTypeEnum.adBSTR);
                //table.Columns.Append("Path", ADOX.DataTypeEnum.adBSTR);
                //table.Columns.Append("LastWriteTime", ADOX.DataTypeEnum.adBSTR);
                try
                {
                    cat.Create(connectionString);


                    ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;


                    con.Execute("CREATE TABLE Files( [Id] Text, [Path] Text, [LastWriteTime] Text)", out _);

                    if (con != null)
                        con.Close();

                    result = true;
                }
                catch (Exception ex)
                {
                    result = false;
                }
                cat = null;
                return result;
            }
        }
        void llog(string @string)
        {
            logobj.Text += @string + Environment.NewLine;
            logobj.ScrollToEnd();
        }
        internal static EnvDTE.Project GetActiveProject()
        {
            EnvDTE.DTE dte = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)) as EnvDTE.DTE;
            EnvDTE.Project activeProject = null;
            string endWith = ((dte.Solution.SolutionBuild as EnvDTE80.SolutionBuild2).StartupProjects as System.Array).GetValue(0).ToString();
            foreach (EnvDTE.Project project in dte.Solution.Projects)
            {
                if (project.FullName != null && project.FullName.EndsWith(endWith))
                {
                    return project;
                }
            }
            return null;
        }
        #endregion
    }
}