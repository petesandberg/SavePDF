namespace SavePDF
{
    using System;
    using System.Collections;
    using System.Runtime.InteropServices;
    using SolidWorks.Interop.sldworks;
    using SolidWorks.Interop.swconst;



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
            this.doc.DestroyNotify += new DDrawingDocEvents_DestroyNotifyEventHandler(this.OnDestroy);
            this.doc.NewSelectionNotify += new DDrawingDocEvents_NewSelectionNotifyEventHandler(this.OnNewSelection);
           
            this.doc.FileSaveNotify += new DDrawingDocEvents_FileSaveNotifyEventHandler(this.doc_FileSaveNotify);

            this.doc.DestroyNotify2  += new DDrawingDocEvents_DestroyNotify2EventHandler(this.doc_FileCloseNotify);
            ConnectModelViews();

            return true;
        }

        private int SavePDFCopy(string fileName)
        {
            try
            {
                ModelDocExtension swModExt = this.ModelDocument.Extension;
                ExportPdfData swExportPDFData = (ExportPdfData)this.SolidworksApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
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


                string pdffilename = fileName.ToUpper().Replace(".SLDDRW", string.Empty);

                if (this.UserAddin.AppendRevision)
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
                    Sheet swSheet = (Sheet)this.doc.GetCurrentSheet();
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

        private int doc_FileSaveNotify(string fileName)
        {
            return this.SavePDFCopy(fileName);
        }
 
        private int doc_FileCloseNotify(int reason)
        {
            if ((reason == (int)swDestroyNotifyType_e.swDestroyNotifyDestroy) && this.UserAddin.SaveOnClose)
            {
                return this.SavePDFCopy(((ModelDoc2)this.doc).GetPathName());
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

}
