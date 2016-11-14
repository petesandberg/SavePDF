// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PropManPageHandler.cs" company="Airgas">
//   Airgas Inc. 2016
// </copyright>
// <summary>
//   Defines the PMPHandler type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace SavePDF
{
    using SolidWorks.Interop.swconst;
    using SolidWorks.Interop.swpublished;

    /// <summary>
    /// The Property Manager Page handler.
    /// </summary>
    public class PropManPageHandler : IPropertyManagerPage2Handler9
    {
        /// <summary>
        /// The user addin.
        /// </summary>
        private readonly SwAddin userAddin;

        //// private ISldWorks solidworksApp;

        public PropManPageHandler(SwAddin addin)
        {
            this.userAddin = addin;
            //// this.solidworksApp = this.userAddin.SwApp;
        }

        // Implement these methods from the interface
        
        /// <summary>
        /// After the page closes.
        /// </summary>
        public void AfterClose()
        {
            // This function must contain code, even if it does nothing, to prevent the
            // .NET runtime environment from doing garbage collection at the wrong time.
            // ReSharper disable once RedundantAssignment
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
                case UserPMPage.RemovePreviousCheckboxId:
                    this.userAddin.ShowPDF = status;
                    break;
                    //case UserPMPage.OnCloseCheckboxId:
                    //    this.userAddin.SaveOnClose = status;
                    //    break;
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

        public bool OnKeystroke(int wparam, int message, int lparam, int id)
        {
            return true;
        }

        public void OnPopupMenuItem(int id)
        {
        }

        public void OnPopupMenuItemUpdate(int id, ref int retval)
        {
        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnSliderPositionChanged(int id, double value)
        {
        }

        public void OnSliderTrackingCompleted(int id, double value)
        {
        }

        public bool OnSubmitSelection(int id, object selection, int selType, ref string itemText)
        {
            return true;
        }

        public bool OnTabClicked(int id)
        {
            return true;
        }

        public void OnUndo()
        {
        }

        public void OnWhatsNew()
        {
        }

        public void OnGainedFocus(int id)
        {
        }

        public void OnListboxRMBUp(int id, int posX, int posY)
        {
        }

        public void OnLostFocus(int id)
        {
        }

        public void OnRedo()
        {
        }

        public int OnWindowFromHandleControlCreated(int id, bool status)
        {
            return 0;
        }
    }
}
