using System.Runtime.InteropServices;
namespace MAPI.NET
{
    partial class MAPISession
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (mapiSpooler_ != null)
                {
                    mapiSpooler_.Dispose();
                    mapiSpooler_ = null;
                }
                if (MsgStore != null)
                {
                    MsgStore.Dispose();
                    MsgStore = null;
                }
                if (session_ != null)
                {
                    Marshal.ReleaseComObject(session_);
                    session_ = null;
                    MAPINative.MAPIUninitialize();
                }
            }
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
        }

        #endregion
    }
}
