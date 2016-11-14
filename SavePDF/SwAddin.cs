// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SwAddin.cs" company="">
//   Airgas 2016
// </copyright>
// <summary>
//   Saves a PDF copy of a drawing.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System.Windows.Forms.VisualStyles;

namespace SavePDF
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    using Microsoft.Win32;

    using SolidWorks.Interop.sldworks;
    using SolidWorks.Interop.swconst;
    using SolidWorks.Interop.swpublished;

    using SolidWorksTools;
    using SolidWorksTools.File;

    using Attribute = System.Attribute;
    using Environment = System.Environment;

    /// <summary>
    ///     Saves a PDF copy of a drawing.
    /// </summary>
    /// [Guid("6dee52ef-da6a-4782-a482-2ecbdb284e98"), ComVisible(true)]
    [Guid("353B343A-427E-4438-97EE-847C6223151A")]
    [ComVisible(true)]
    [SwAddin(Description = "Saves a PDF when you save a drawing", Title = "SavePDF", LoadAtStartup = true)]
    public class SwAddin : ISwAddin
    {
        #region Constants

        /// <summary>
        /// The main cmd group id.
        /// </summary>
        public const int MainCmdGroupId = 1;

        /// <summary>
        /// The main item i d 1.
        /// </summary>
        public const int MainItemId1 = 0;

        /// <summary>
        /// The description value name.
        /// </summary>
        private const string DescriptionValueName = "AppendDescription";

        /// <summary>
        /// The location value name.
        /// </summary>
        private const string LocationValueName = "PDFLocation";

        /// <summary>
        /// The options key name.
        /// </summary>
        private const string OptionsKeyName = "Software\\Airgas Inc\\SavePDF";

        /// <summary>
        /// The revision value name.
        /// </summary>
        private const string RevisionValueName = "AppendRevision";

        /// <summary>
        /// The show pdf value name.
        /// </summary>
        private const string ShowPdfValueName = "ShowPDF";

        /// <summary>
        /// The show pdf value name.
        /// </summary>
        private const string RemovePrevValueName = "RemovePrevious";

        /// <summary>
        /// The save on close value name.
        /// </summary>
        private const string SaveOnCloseValueName = "SaveOnClose";

        #endregion

        #region Fields

        /// <summary>
        /// The sw event ptr.
        /// </summary>
        private SldWorks solidworksEventPtr;

        /// <summary>
        /// The addin id.
        /// </summary>
        private int addinId;

        /// <summary>
        /// The i bmp.
        /// </summary>
        private BitmapHandler iBmp;

        /// <summary>
        /// The open docs.
        /// </summary>
        private Hashtable openDocs = new Hashtable();

        /// <summary>
        /// The ppage.
        /// </summary>
        private UserPMPage ppage;

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether append description.
        /// </summary>
        public bool AppendDescription { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether append revision.
        /// </summary>
        public bool AppendRevision { get; set; }

        // Public Properties

        /// <summary>
        /// Gets the cmd mgr.
        /// </summary>
        public ICommandManager CmdMgr { get; private set; }

        /// <summary>
        /// Gets the open docs.
        /// </summary>
        public Hashtable OpenDocs
        {
            get
            {
                return this.openDocs;
            }
        }

        /// <summary>
        /// Gets or sets the pdf location.
        /// </summary>
        public string PDFLocation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether show pdf.
        /// </summary>
        public bool ShowPDF { get; set; }

        /// <summary>
        /// Gets the sw app.
        /// </summary>
        public ISldWorks SwApp { get; private set; }

        public bool SaveOnClose { get; set; }

        public bool RemovePrevious { get; set; }
        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The register function.
        /// </summary>
        /// <param name="t">
        /// The t.
        /// </param>
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            Type type = typeof(SwAddin);

            SwAddinAttribute solidworksAttribute = type.GetCustomAttributes(false).OfType<SwAddinAttribute>().Select(attr => attr).FirstOrDefault();

            try
            {
                RegistryKey hklm = Registry.LocalMachine;
                RegistryKey hkcu = Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID + "}";
                RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", solidworksAttribute.Description);
                addinkey.SetValue("Title", solidworksAttribute.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(solidworksAttribute.LoadAtStartup), RegistryValueKind.DWord);
            }
            catch (NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

                MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        /// <summary>
        /// The unregister function.
        /// </summary>
        /// <param name="t">
        /// The t.
        /// </param>
        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                RegistryKey hklm = Registry.LocalMachine;
                RegistryKey hkcu = Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        /// <summary>
        /// The add command mgr.
        /// </summary>
        public void AddCommandMgr()
        {
            if (this.iBmp == null)
            {
                this.iBmp = new BitmapHandler();
            }

            const string Title = "Save PDF";
            const string ToolTip = "Save PDF";

            int[] docTypes =
                {
                    (int)swDocumentTypes_e.swDocASSEMBLY, (int)swDocumentTypes_e.swDocDRAWING, 
                    (int)swDocumentTypes_e.swDocPART
                };

            Assembly thisAssembly = Assembly.GetAssembly(this.GetType());

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;

            // get the ID information stored in the registry
            bool getDataResult = this.CmdMgr.GetGroupDataFromRegistry(MainCmdGroupId, out registryIDs);

            int[] knownIDs = { MainItemId1 };

            if (getDataResult)
            {
                if (!this.CompareIDs((int[])registryIDs, knownIDs))
                {
                    // if the IDs don't match, reset the commandGroup
                    ignorePrevious = true;
                }
            }

            ICommandGroup cmdGroup = this.CmdMgr.CreateCommandGroup2(
                MainCmdGroupId, 
                Title, 
                ToolTip, 
                "", 
                -1, 
                ignorePrevious, 
                ref cmdGroupErr);

            cmdGroup.LargeIconList = this.iBmp.CreateFileFromResourceBitmap("SavePDF.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = this.iBmp.CreateFileFromResourceBitmap("SavePDF.ToolbarSmall.bmp", thisAssembly);

            const int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            int cmdIndex0 = cmdGroup.AddCommandItem2(
                "Save PDF", 
                -1, 
                "Save PDF Options", 
                "Save PDF Options", 
                0, 
                "ShowPMP", 
                "EnablePMP", 
                MainItemId1, 
                menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            foreach (int type in docTypes)
            {
                CommandTab cmdTab = this.CmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)
                {
                    // if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                    bool res = this.CmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                // if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = this.CmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    var cmdIDs = new int[1];
                    var textType = new int[1];

                    cmdIDs[0] = cmdGroup.CommandID[cmdIndex0];

                    textType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdBox.AddCommands(cmdIDs, textType);
                }
            }

            thisAssembly = null;
        }

        /// <summary>
        /// The add pmp.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool AddPMP()
        {
            this.ppage = new UserPMPage(this);
            return true;
        }

        /// <summary>
        /// The attach event handlers.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool AttachEventHandlers()
        {
            this.AttachSwEvents();

            // Listen for events on all currently open docs
            this.AttachEventsToAllDocuments();
            return true;
        }

        /// <summary>
        /// The attach events to all documents.
        /// </summary>
        public void AttachEventsToAllDocuments()
        {
            var modDoc = (ModelDoc2)this.SwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!this.openDocs.Contains(modDoc))
                {
                    this.AttachModelDocEventHandler(modDoc);
                }
                else if (this.openDocs.Contains(modDoc))
                {
                    bool connected = false;
                    var docHandler = (DocumentEventHandler)this.openDocs[modDoc];
                    if (docHandler != null)
                    {
                        connected = docHandler.ConnectModelViews();
                    }
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        /// <summary>
        /// The attach model doc event handler.
        /// </summary>
        /// <param name="modDoc">
        /// The mod doc.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
            {
                return false;
            }

            DocumentEventHandler docHandler = null;

            if (!this.openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }

                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }

                    default:
                        {
                            return false; // Unsupported document type
                        }
                }

                docHandler.AttachEventHandlers();
                this.openDocs.Add(modDoc, docHandler);
            }

            return true;
        }

        /// <summary>
        /// The compare i ds.
        /// </summary>
        /// <param name="storedIDs">
        /// The stored i ds.
        /// </param>
        /// <param name="addinIDs">
        /// The addin i ds.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            var storedList = new List<int>(storedIDs);
            var addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }

            for (int i = 0; i < addinList.Count; i++)
            {
                if (addinList[i] != storedList[i])
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// The connect to sw.
        /// </summary>
        /// <param name="ThisSW">
        /// The this sw.
        /// </param>
        /// <param name="cookie">
        /// The cookie.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool ConnectToSW(object ThisSW, int cookie)
        {
            this.SwApp = (ISldWorks)ThisSW;
            this.addinId = cookie;

            // Setup callbacks
            this.SwApp.SetAddinCallbackInfo(0, this, this.addinId);

            

            this.CmdMgr = this.SwApp.GetCommandManager(cookie);
            this.AddCommandMgr();

            

            #region Setup the Event Handlers

            this.solidworksEventPtr = (SldWorks)this.SwApp;
            this.openDocs = new Hashtable();
            this.AttachEventHandlers();

            #endregion

            this.ReadOptions();

            #region Setup Property Manager

            this.AddPMP();

            #endregion

            return true;
        }

        /// <summary>
        /// The detach event handlers.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool DetachEventHandlers()
        {
            this.DetachSwEvents();

            // Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = this.openDocs.Count;
            var keys = new object[numKeys];

            // Remove all document event handlers
            this.openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)this.openDocs[key];
                docHandler.DetachEventHandlers(); // This also removes the pair from the hash
                docHandler = null;
            }

            return true;
        }

        /// <summary>
        /// The detach model event handler.
        /// </summary>
        /// <param name="modDoc">
        /// The mod doc.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)this.openDocs[modDoc];
            this.openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        /// <summary>
        /// The disconnect from sw.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool DisconnectFromSW()
        {
            this.RemoveCommandMgr();
            this.RemovePMP();
            this.DetachEventHandlers();

            Marshal.ReleaseComObject(this.CmdMgr);
            this.CmdMgr = null;
            Marshal.ReleaseComObject(this.SwApp);
            this.SwApp = null;

            // The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        /// <summary>
        /// The enable pmp.
        /// </summary>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int EnablePMP()
        {
            if (this.SwApp.ActiveDoc != null)
            {
                return 1;
            }

            return 0;
        }

        // Events
        
        /// <summary>
        /// The on doc change.
        /// </summary>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int OnDocChange()
        {
            return 0;
        }

        /// <summary>
        /// The on doc load.
        /// </summary>
        /// <param name="docTitle">
        /// The doc title.
        /// </param>
        /// <param name="docPath">
        /// The doc path.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        /// <summary>
        /// The on file new.
        /// </summary>
        /// <param name="newDoc">
        /// The new doc.
        /// </param>
        /// <param name="docType">
        /// The doc type.
        /// </param>
        /// <param name="templateName">
        /// The template name.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            this.AttachEventsToAllDocuments();
            return 0;
        }

        /// <summary>
        /// The on model change.
        /// </summary>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int OnModelChange()
        {
            return 0;
        }

        /// <summary>
        /// The read options.
        /// </summary>
        public void ReadOptions()
        {
            try
            {
                RegistryKey hkcu = Registry.CurrentUser;
                RegistryKey optionsKey = hkcu.OpenSubKey(OptionsKeyName, true) ?? hkcu.CreateSubKey(OptionsKeyName);

                this.PDFLocation =
                    (string)
                    optionsKey.GetValue(
                        LocationValueName, 
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                this.AppendRevision = Convert.ToBoolean(optionsKey.GetValue(RevisionValueName, true));
                this.AppendDescription = Convert.ToBoolean(optionsKey.GetValue(DescriptionValueName, true));
                this.ShowPDF = Convert.ToBoolean(optionsKey.GetValue(ShowPdfValueName, false));
                this.RemovePrevious = Convert.ToBoolean(optionsKey.GetValue(RemovePrevValueName, false));
                //this.SaveOnClose = Convert.ToBoolean(optionsKey.GetValue(SaveOnCloseValueName, false));
                this.SaveOnClose = true;
            }
            catch (Exception nl)
            {
                MessageBox.Show("There was a problem reading the options: \n\"" + nl.Message + "\"");
            }
        }

        /// <summary>
        /// The remove command mgr.
        /// </summary>
        public void RemoveCommandMgr()
        {
            this.iBmp.Dispose();

            this.CmdMgr.RemoveCommandGroup(MainCmdGroupId);
        }

        /// <summary>
        /// The remove pmp.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool RemovePMP()
        {
            this.ppage = null;
            return true;
        }

        /// <summary>
        /// The show pmp.
        /// </summary>
        public void ShowPMP()
        {
            if (this.ppage != null)
            {
                this.ppage.Show();
            }
        }

        /// <summary>
        /// The write options.
        /// </summary>
        public void WriteOptions()
        {
            try
            {
                RegistryKey hkcu = Registry.CurrentUser;

                RegistryKey optionsKey = hkcu.OpenSubKey(OptionsKeyName, true) ?? hkcu.CreateSubKey(OptionsKeyName);

                optionsKey.SetValue(LocationValueName, this.PDFLocation, RegistryValueKind.String);
                optionsKey.SetValue(RevisionValueName, this.AppendRevision, RegistryValueKind.DWord);
                optionsKey.SetValue(DescriptionValueName, this.AppendDescription, RegistryValueKind.DWord);
                optionsKey.SetValue(ShowPdfValueName, this.ShowPDF, RegistryValueKind.DWord);
                optionsKey.SetValue(SaveOnCloseValueName, this.SaveOnClose, RegistryValueKind.DWord);
                optionsKey.SetValue(RemovePrevValueName, this.RemovePrevious, RegistryValueKind.DWord);
            }
            catch (Exception nl)
            {
                MessageBox.Show("There was a problem writing the options: \n\"" + nl.Message + "\"");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// The attach sw events.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool AttachSwEvents()
        {
            try
            {
                this.solidworksEventPtr.ActiveDocChangeNotify += this.OnDocChange;
                this.solidworksEventPtr.DocumentLoadNotify2 += this.OnDocLoad;
                this.solidworksEventPtr.FileNewNotify2 += this.OnFileNew;
                this.solidworksEventPtr.ActiveModelDocChangeNotify += this.OnModelChange;
                this.solidworksEventPtr.FileOpenPostNotify += this.FileOpenPostNotify;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// The detach sw events.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool DetachSwEvents()
        {
            try
            {
                this.solidworksEventPtr.ActiveDocChangeNotify -= this.OnDocChange;
                this.solidworksEventPtr.DocumentLoadNotify2 -= this.OnDocLoad;
                this.solidworksEventPtr.FileNewNotify2 -= this.OnFileNew;
                this.solidworksEventPtr.ActiveModelDocChangeNotify -= this.OnModelChange;
                this.solidworksEventPtr.FileOpenPostNotify -= this.FileOpenPostNotify;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// The file open post notify.
        /// </summary>
        /// <param name="FileName">
        /// The file name.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        private int FileOpenPostNotify(string FileName)
        {
            this.AttachEventsToAllDocuments();
            return 0;
        }

        #endregion
    }
}