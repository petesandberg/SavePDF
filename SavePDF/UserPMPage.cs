
namespace SavePDF
{

    using System;
    using SolidWorks.Interop.sldworks;
    using SolidWorks.Interop.swconst;

    /// <summary>
    /// The user pm page.
    /// </summary>
    public class UserPMPage
    {
        // Local Objects
        private IPropertyManagerPage2 swPropertyPage;
        private PropManPageHandler handler;
        private readonly ISldWorks iSwApp;
        private readonly SwAddin userAddin;

        #region Property Manager Page Controls
        //Groups
        IPropertyManagerPageGroup SavePDFGroup;

        //Controls
        IPropertyManagerPageTextbox LocationTextbox;
        IPropertyManagerPageCheckbox revisionCheckbox;
        IPropertyManagerPageCheckbox descriptionCheckbox;
        IPropertyManagerPageCheckbox showCheckbox;
        //IPropertyManagerPageCheckbox ocCloseCheckbox;

        //Control IDs
        public const int group1ID = 0;

        public const int LocationTextboxId = 1;
        public const int RevisionCheckboxId = 2;
        public const int DescriptionCheckboxId = 3;
        public const int ShowCheckboxId = 4;
        //public const int OnCloseCheckboxId = 5;
        #endregion

        public UserPMPage(SwAddin addin)
        {
            this.userAddin = addin;
            if (this.userAddin != null)
            {
                this.iSwApp = this.userAddin.SwApp;
                this.CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwAddin not set.");
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            const int Options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

            this.handler = new PropManPageHandler(this.userAddin);
            this.swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Save PDF", Options, this.handler, ref errors);
            if (this.swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    this.AddControls();
                }
                catch (Exception e)
                {
                    this.iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }


        // Controls are displayed on the page top to bottom in the order 
        // in which they are added to the object.
        protected void AddControls()
        {
            short controlType = -1;
            short align = -1;
            int options = -1;


            // Add the groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            this.SavePDFGroup = (IPropertyManagerPageGroup)this.swPropertyPage.AddGroupBox(group1ID, "Save PDF", options);

            // Add the controls to group1

            // LocationTextbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            this.LocationTextbox = (IPropertyManagerPageTextbox)this.SavePDFGroup.AddControl(LocationTextboxId, controlType, "PDF Location", align, options, "The File location where the PDFs will be saved to.");

            // Rev Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            this.revisionCheckbox = (IPropertyManagerPageCheckbox)this.SavePDFGroup.AddControl(RevisionCheckboxId, controlType, "Append Revision", align, options, "Append the Revision to the filename");
            
            // Description Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            this.descriptionCheckbox = (IPropertyManagerPageCheckbox)this.SavePDFGroup.AddControl(DescriptionCheckboxId, controlType, "Append Description", align, options, "Append the Description to the filename");

            // Show Checkbox
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            this.showCheckbox = (IPropertyManagerPageCheckbox)SavePDFGroup.AddControl(ShowCheckboxId, controlType, "Show PDF", align, options, "Show the PDF after saving");
 
            // On Close Checkbox
            //controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            //align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            //options = (int)swAddControlOptions_e.swControlOptions_Enabled |
            //          (int)swAddControlOptions_e.swControlOptions_Visible;

            //this.ocCloseCheckbox = (IPropertyManagerPageCheckbox)this.SavePDFGroup.AddControl(OnCloseCheckboxId, controlType, "Save on Close", align, options, "Save the PDF on close");
        }

        public void Show()
        {
            if (this.swPropertyPage != null)
            {
                this.LocationTextbox.Text = this.userAddin.PDFLocation;
                this.revisionCheckbox.Checked = this.userAddin.AppendRevision;
                this.descriptionCheckbox.Checked = this.userAddin.AppendDescription;
                this.showCheckbox.Checked = this.userAddin.ShowPDF;
                //this.ocCloseCheckbox.Checked = this.userAddin.SaveOnClose;
                this.swPropertyPage.Show();
            }
        }
    }
}
