namespace SavePDF
{
    using System;
    using System.Collections;
    using System.Runtime.InteropServices;
    using SolidWorks.Interop.sldworks;
    using SolidWorks.Interop.swconst;

    public class DocumentEventHandler
    {
        protected ISldWorks SolidworksApp { get; set; }

        protected ModelDoc2 ModelDocument { get; set; }
        
        protected SwAddin UserAddin { get; set; }

        protected Hashtable OpenModelViews { get; set; }

        public DocumentEventHandler(ModelDoc2 modDoc, SwAddin addin)
        {
            ModelDocument = modDoc;
            UserAddin = addin;
            SolidworksApp = (ISldWorks)UserAddin.SwApp;
            OpenModelViews = new Hashtable();
        }

        virtual public bool AttachEventHandlers()
        {
            return true;
        }

        virtual public bool DetachEventHandlers()
        {
            return true;
        }

        public bool ConnectModelViews()
        {
            IModelView mView;
            mView = (IModelView)ModelDocument.GetFirstModelView();

            while (mView != null)
            {
                if (!OpenModelViews.Contains(mView))
                {
                    DocView dView = new DocView(UserAddin, mView, this);
                    dView.AttachEventHandlers();
                    OpenModelViews.Add(mView, dView);
                }
                mView = (IModelView)mView.GetNext();
            }
            return true;
        }

        public bool DisconnectModelViews()
        {
            //Close events on all currently open docs
            DocView dView;
            int numKeys;
            numKeys = OpenModelViews.Count;

            if (numKeys == 0)
            {
                return false;
            }


            object[] keys = new object[numKeys];

            // Remove all ModelView event handlers
            OpenModelViews.Keys.CopyTo(keys, 0);
            foreach (ModelView key in keys)
            {
                dView = (DocView)OpenModelViews[key];
                dView.DetachEventHandlers();
                OpenModelViews.Remove(key);
                dView = null;
            }
            return true;
        }

        public bool DetachModelViewEventHandler(ModelView mView)
        {
            DocView dView;
            if (OpenModelViews.Contains(mView))
            {
                dView = (DocView)OpenModelViews[mView];
                OpenModelViews.Remove(mView);
                mView = null;
                dView = null;
            }
            return true;
        }
    }

    public class PartEventHandler : DocumentEventHandler
    {
        PartDoc doc;

        public PartEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            doc = (PartDoc)ModelDocument;
        }

        override public bool AttachEventHandlers()
        {
            doc.DestroyNotify += new DPartDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify += new DPartDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);

            ConnectModelViews();

            return true;
        }

        override public bool DetachEventHandlers()
        {
            doc.DestroyNotify -= new DPartDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify -= new DPartDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);

            DisconnectModelViews();

            UserAddin.DetachModelEventHandler(ModelDocument);
            return true;
        }

        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }
    }

    public class AssemblyEventHandler : DocumentEventHandler
    {
        AssemblyDoc doc;
        SwAddin swAddin;

        public AssemblyEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            doc = (AssemblyDoc)ModelDocument;
            swAddin = addin;
        }

        override public bool AttachEventHandlers()
        {
            doc.DestroyNotify += new DAssemblyDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify += new DAssemblyDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            doc.ComponentStateChangeNotify2 += new DAssemblyDocEvents_ComponentStateChangeNotify2EventHandler(ComponentStateChangeNotify2);
            doc.ComponentStateChangeNotify += new DAssemblyDocEvents_ComponentStateChangeNotifyEventHandler(ComponentStateChangeNotify);
            doc.ComponentVisualPropertiesChangeNotify += new DAssemblyDocEvents_ComponentVisualPropertiesChangeNotifyEventHandler(ComponentVisualPropertiesChangeNotify);
            doc.ComponentDisplayStateChangeNotify += new DAssemblyDocEvents_ComponentDisplayStateChangeNotifyEventHandler(ComponentDisplayStateChangeNotify);
            ConnectModelViews();

            return true;
        }

        override public bool DetachEventHandlers()
        {
            doc.DestroyNotify -= new DAssemblyDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify -= new DAssemblyDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            doc.ComponentStateChangeNotify2 -= new DAssemblyDocEvents_ComponentStateChangeNotify2EventHandler(ComponentStateChangeNotify2);
            doc.ComponentStateChangeNotify -= new DAssemblyDocEvents_ComponentStateChangeNotifyEventHandler(ComponentStateChangeNotify);
            doc.ComponentVisualPropertiesChangeNotify -= new DAssemblyDocEvents_ComponentVisualPropertiesChangeNotifyEventHandler(ComponentVisualPropertiesChangeNotify);
            doc.ComponentDisplayStateChangeNotify -= new DAssemblyDocEvents_ComponentDisplayStateChangeNotifyEventHandler(ComponentDisplayStateChangeNotify);
            DisconnectModelViews();

            UserAddin.DetachModelEventHandler(ModelDocument);
            return true;
        }

        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }

        //attach events to a component if it becomes resolved
        protected int ComponentStateChange(object componentModel, short newCompState)
        {
            ModelDoc2 modDoc = (ModelDoc2)componentModel;
            swComponentSuppressionState_e newState = (swComponentSuppressionState_e)newCompState;


            switch (newState)
            {

                case swComponentSuppressionState_e.swComponentFullyResolved:
                    {
                        if ((modDoc != null) & !this.swAddin.OpenDocs.Contains(modDoc))
                        {
                            this.swAddin.AttachModelDocEventHandler(modDoc);
                        }
                        break;
                    }

                case swComponentSuppressionState_e.swComponentResolved:
                    {
                        if ((modDoc != null) & !this.swAddin.OpenDocs.Contains(modDoc))
                        {
                            this.swAddin.AttachModelDocEventHandler(modDoc);
                        }
                        break;
                    }

            }
            return 0;
        }

        protected int ComponentStateChange(object componentModel)
        {
            ComponentStateChange(componentModel, (short)swComponentSuppressionState_e.swComponentResolved);
            return 0;
        }


        public int ComponentStateChangeNotify2(object componentModel, string CompName, short oldCompState, short newCompState)
        {
            return ComponentStateChange(componentModel, newCompState);
        }

        int ComponentStateChangeNotify(object componentModel, short oldCompState, short newCompState)
        {
            return ComponentStateChange(componentModel, newCompState);
        }

        int ComponentDisplayStateChangeNotify(object swObject)
        {
            Component2 component = (Component2)swObject;
            ModelDoc2 modDoc = (ModelDoc2)component.GetModelDoc();

            return ComponentStateChange(modDoc);
        }

        int ComponentVisualPropertiesChangeNotify(object swObject)
        {
            Component2 component = (Component2)swObject;
            ModelDoc2 modDoc = (ModelDoc2)component.GetModelDoc();

            return ComponentStateChange(modDoc);
        }




    }

    public class DrawingEventHandler : DocumentEventHandler
    {
        DrawingDoc doc;

        public DrawingEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            doc = (DrawingDoc)ModelDocument;
        }

        override public bool AttachEventHandlers()
        {
            doc.DestroyNotify += new DDrawingDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify += new DDrawingDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            doc.FileSaveNotify += new DDrawingDocEvents_FileSaveNotifyEventHandler(this.doc_FileSaveNotify);

            ConnectModelViews();

            return true;
        }

        private int doc_FileSaveNotify(string FileName)
        {
            try
            {
                ModelDocExtension swModExt = this.ModelDocument.Extension;
                ExportPdfData swExportPDFData = (ExportPdfData)this.SolidworksApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                Sheet swSheet;
                bool boolstatus = false;
                int errors = 0;
                int warnings = 0;
                string[] obj = null;

                string description;
                string resolvedDescription;
                bool wasResolved;
                string revision;
                string resolverdRevision;
                CustomPropertyManager swCustProp = swModExt.CustomPropertyManager[string.Empty];

                swCustProp.Get5("Description", false, out description, out resolvedDescription, out wasResolved);
                swCustProp.Get5("Revision", false, out revision, out resolverdRevision, out wasResolved);


                string pdffilename = FileName.ToUpper().Replace(".SLDDRW", string.Empty);

                if (this.userAddin.AppendRevision)
                {
                    pdffilename += ", Rev. " + revision;
                }
              
                if (this.UserAddin.AppendDescription)
                {
                    pdffilename += ", " + description;
                       
                }


                pdffilename = pdffilename.Substring(pdffilename.LastIndexOf('\\'));
                pdffilename = this.UserAddin.PDFLocation + pdffilename;
                pdffilename = pdffilename + ".pdf";

                obj = this.doc.GetSheetNames();
                int count = 0;
                count = obj.Length;
                int i = 0;
                object[] objs = new object[count - 1];
                DispatchWrapper[] arrObjIn = new DispatchWrapper[count - 1];

                // Activate each drawing sheet, except the last drawing sheet, for 
                // demonstration purposes only and add each sheet to an array 
                // of drawing sheets 
                for (i = 0; i < count - 1; i++)
                {
                    boolstatus = this.doc.ActivateSheet(obj[i]);
                    swSheet = (Sheet)this.doc.GetCurrentSheet();
                    objs[i] = swSheet;
                    arrObjIn[i] = new DispatchWrapper(objs[i]);


                }

                // Save the drawings sheets to a PDF file 
                swExportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, arrObjIn);
                swExportPDFData.ViewPdfAfterSaving = this.UserAddin.ShowPDF;

                swModExt.SaveAs(pdffilename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Problem Saving the PDF. " + e.Message);
            }

            return 0;
        }

        override public bool DetachEventHandlers()
        {
            doc.DestroyNotify -= new DDrawingDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify -= new DDrawingDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);

            doc.FileSaveNotify -= new DDrawingDocEvents_FileSaveNotifyEventHandler(this.doc_FileSaveNotify);

            DisconnectModelViews();

            UserAddin.DetachModelEventHandler(ModelDocument);
            return true;
        }

        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }
    }

    public class DocView
    {
        ISldWorks iSwApp;
        SwAddin userAddin;
        ModelView mView;
        DocumentEventHandler parent;

        public DocView(SwAddin addin, IModelView mv, DocumentEventHandler doc)
        {
            userAddin = addin;
            mView = (ModelView)mv;
            iSwApp = (ISldWorks)userAddin.SwApp;
            parent = doc;
        }

        public bool AttachEventHandlers()
        {
            mView.DestroyNotify2 += new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);
            mView.RepaintNotify += new DModelViewEvents_RepaintNotifyEventHandler(OnRepaint);
            return true;
        }

        public bool DetachEventHandlers()
        {
            mView.DestroyNotify2 -= new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);
            mView.RepaintNotify -= new DModelViewEvents_RepaintNotifyEventHandler(OnRepaint);
            parent.DetachModelViewEventHandler(mView);
            return true;
        }

        //EventHandlers
        public int OnDestroy(int destroyType)
        {
            switch (destroyType)
            {
                case (int)swDestroyNotifyType_e.swDestroyNotifyHidden:
                    return 0;

                case (int)swDestroyNotifyType_e.swDestroyNotifyDestroy:
                    return 0;
            }

            return 0;
        }

        public int OnRepaint(int repaintType)
        {
            return 0;
        }
    }

}
