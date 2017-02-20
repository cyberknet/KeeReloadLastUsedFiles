using System;
using System.Collections.Generic;
using System.Text;
using KeePass.Plugins;
using KeePass;
using KeePass.Util.XmlSerialization;
using KeePassLib.Serialization;
using System.Xml;
using System.Windows.Forms;
using System.IO;
using KeePass.Forms;

namespace KeeReloadLastUsedFiles
{
    /// <summary>
    ///     Plugin to enable KeePass to save and reload last open key databases.
    /// </summary>
    /// <remarks>KeePass SDK documentation: http://keepass.info/help/v2_dev/plg_index.html</remarks>
    public class KeeReloadLastUsedFilesExt : Plugin
    {
        private IPluginHost pluginHost = null;
        private const string configKey = "ReloadLUF.LastUsedFiles";
        private bool exiting = false;
        private Object lockObject = new Object();

        /// <summary>
        ///     Returns the URL where KeePass can check for updates of this plugin
        /// </summary>
        public override string UpdateUrl
        {
            get { return @"https://raw.githubusercontent.com/cyberknet/KeeReloadLastUsedFiles/master/latest.txt"; }
        }

        /// <summary>
        ///     Called when the Plugin is being loaded which happens on startup of KeePass
        /// </summary>
        /// <returns>True if the plugin loaded successfully, false if not</returns>
        public override bool Initialize(IPluginHost host)
        {
            pluginHost = host;
            pluginHost.MainWindow.FileClosingPre += MainWindow_FileClosingPre;
            pluginHost.MainWindow.FormLoadPost += MainWindow_FormLoadPost;
            return true; 
        }

        private void MainWindow_FormLoadPost(object sender, EventArgs e)
        {
            try
            {
                string xml = pluginHost.CustomConfig.GetString(configKey);
                if (!string.IsNullOrEmpty(xml))
                {
                    var ser = new XmlSerializerEx(typeof(IOConnectionInfo[]));
                    var stream = GenerateStreamFromString(xml);
                    var connections = (IOConnectionInfo[]) ser.Deserialize(stream);
                    // loop through all the saved connections
                    foreach (var ioLastFile in connections)
                    {
                        bool isLoaded = false;
                        var cmpPath = ioLastFile.Path.ToLower().Trim();

                        foreach(var loadedDoc in pluginHost.MainWindow.DocumentManager.Documents)
                        {
                            if (loadedDoc.LockedIoc != null && loadedDoc.LockedIoc.Path.Length > 0)
                            {
                                isLoaded = loadedDoc.LockedIoc.Path.ToLower().Trim() == cmpPath;
                            }
                            else if (loadedDoc.Database != null && loadedDoc.Database.IOConnectionInfo.Path.Length > 0)
                            {
                                isLoaded = loadedDoc.Database.IOConnectionInfo.Path.ToLower().Trim() == cmpPath;
                            }
                            if (isLoaded)
                                break;
                        }
                        if (!isLoaded)
                            pluginHost.MainWindow.OpenDatabase(ioLastFile, null, false);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to load list of previously open password database files from config file:\n\n" + ex.ToString());
            }
        }

        private void MainWindow_FileClosingPre(object sender, KeePass.Forms.FileClosingEventArgs e)
        {
            lock(lockObject)
            {
                if (((e.Flags & FileEventFlags.Exiting) == FileEventFlags.Exiting) && !exiting)
                {
                    exiting = true;
                    string xml = string.Empty;
                    try
                    {
                        var ser = new XmlSerializerEx(typeof(IOConnectionInfo[]));
                        var builder = new StringBuilder();
                        XmlWriter writer = XmlWriter.Create(builder);

                        var connections = new List<IOConnectionInfo>();
                        foreach (var document in pluginHost.MainWindow.DocumentManager.Documents)
                        {
                            if (document.LockedIoc != null && document.LockedIoc.Path.Length > 0)
                                connections.Add(document.LockedIoc);
                            else if (document.Database.IOConnectionInfo != null && document.Database.IOConnectionInfo.Path.Length > 0)
                                connections.Add(document.Database.IOConnectionInfo);
                        }
                        var connectionArray = connections.ToArray();
                        ser.Serialize(writer, connectionArray);
                        writer.Flush();
                        xml = builder.ToString();
                        pluginHost.CustomConfig.SetString(configKey, xml);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Unable to save list of open password database files to config file:\n\n" + ex.ToString());
                    }
                }
            }
        }

        public override void Terminate()
        {
            base.Terminate();
        }

        private Stream GenerateStreamFromString(string s)
        {
            var b = Encoding.Unicode.GetBytes(s);
            var stream = new MemoryStream(b);
            return stream;
        }
    }
}
