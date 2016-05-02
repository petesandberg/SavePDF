namespace SavePDF
{
    using System.Collections;

    using SolidWorks.Interop.sldworks;

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

        public virtual bool AttachEventHandlers()
        {
            return true;
        }

        public virtual bool DetachEventHandlers()
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
}
