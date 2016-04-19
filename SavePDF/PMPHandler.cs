// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PMPHandler.cs" company="Airgas">
//   Airgas Inc. 2016
// </copyright>
// <summary>
//   Defines the PMPHandler type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace SavePDF
{
    using System.IO;
    using System.Windows.Forms;

    using SolidWorks.Interop.sldworks;
    using SolidWorks.Interop.swconst;
    using SolidWorks.Interop.swpublished;

    /// <summary>
    /// The pmp handler.
    /// </summary>
    public class PMPHandler : IPropertyManagerPage2Handler9
    {
        /// <summary>
        /// The user addin.
        /// </summary>
        private readonly SwAddin userAddin;

        /// <summary>
        /// The SolidWorks App.
        /// </summary>
        private ISldWorks solidworksApp;

        public PMPHandler(SwAddin addin)
        {
            this.userAddin = addin;
            this.solidworksApp = this.userAddin.SwApp;
        }

        // Implement these methods from the interface
        
        /// <summary>
        /// After the page closes.
        /// </summary>
        public void AfterClose()
        {
            // This function must contain code, even if it does nothing, to prevent the
            // .NET runtime environment from doing garbage collection at the wrong time.
            int indentSize = System.Diagnostics.Debug.IndentSize;
            System.Diagnostics.Debug.WriteLine(indentSize);
        }

        /// <summary>
        /// The on checkbox check.
        /// </summary>
        /// <param name="id">
        /// The id of the checkbox that was checked.
        /// </param>
        /// <param name="status">
        /// The new value of the checkbox.
        /// </param>
        public void OnCheckboxCheck(int id, bool status)
        {
            switch (id)
            {
                case UserPMPage.DescriptionCheckboxId:
                    this.userAddin.AppendDescription = status;
                    break;
                case UserPMPage.RevisionCheckboxId:
                    this.userAddin.AppendRevision = status;
                    break;
                case UserPMPage.ShowCheckboxId:
                    this.userAddin.ShowPDF = status;
                    break;
            }
        }

        /// <summary>
        /// The on textbox changed.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="text">
        /// The text.
        /// </param>
        public void OnTextboxChanged(int id, string text)
        {
            if (id == UserPMPage.LocationTextboxId)
            {
                this.userAddin.PDFLocation = text;
            }
        }

        /// <summary>
        /// Before the page closes.
        /// </summary>
        /// <param name="reason">
        /// The reason for closure.
        /// </param>
        public void OnClose(int reason)
        {
            // This function must contain code, even if it does nothing, to prevent the
            // .NET runtime environment from doing garbage collection at the wrong time.
            switch (reason)
            {
                case (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay:
                    this.userAddin.WriteOptions();
                    break;

                // case (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel:
                default:
                    this.userAddin.ReadOptions();
                    break;
            }
        }

        public void OnComboboxEditChanged(int id, string text)
        {
        }

        public int OnActiveXControlCreated(int id, bool status)
        {
            return -1;
        }

        public void OnButtonPress(int id)
        {

        }

        public void OnComboboxSelectionChanged(int id, int item)
        {

        }

        public void OnGroupCheck(int id, bool status)
        {
        }

        public void OnGroupExpand(int id, bool status)
        {

        }

        public bool OnHelp()
        {
            return true;
        }

        public void OnListboxSelectionChanged(int id, int item)
        {

        }

        public bool OnNextPage()
        {
            return true;
        }

        public void OnNumberboxChanged(int id, double val)
        {

        }

        public void OnNumberBoxTrackingCompleted(int id, double val)
        {

        }

        public void OnOptionCheck(int id)
        {

        }

        public bool OnPreviousPage()
        {
            return true;
        }

        public void OnSelectionboxCalloutCreated(int id)
        {

        }

        public void OnSelectionboxCalloutDestroyed(int id)
        {

        }

        public void OnSelectionboxFocusChanged(int id)
        {

        }

        public void OnSelectionboxListChanged(int id, int item)
        {

        }

        public void AfterActivation()
        {

        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            return true;
        }

        public void OnPopupMenuItem(int Id)
        {

        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {

        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {

        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {

        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            return true;
        }

        public bool OnTabClicked(int Id)
        {
            return true;
        }

        public void OnUndo()
        {

        }

        public void OnWhatsNew()
        {

        }


        public void OnGainedFocus(int Id)
        {

        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {

        }

        public void OnLostFocus(int Id)
        {

        }

        public void OnRedo()
        {

        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            return 0;
        }
    }
}
