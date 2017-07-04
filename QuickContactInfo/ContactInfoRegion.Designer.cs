namespace QuickContactInfo
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class ContactInfoRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public ContactInfoRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
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
            this.tblContacts = new System.Windows.Forms.TableLayoutPanel();
            this.SuspendLayout();
            // 
            // tblContacts
            // 
            this.tblContacts.AutoScroll = true;
            this.tblContacts.AutoSize = true;
            this.tblContacts.ColumnCount = 4;
            this.tblContacts.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tblContacts.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tblContacts.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tblContacts.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tblContacts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblContacts.Location = new System.Drawing.Point(0, 0);
            this.tblContacts.Name = "tblContacts";
            this.tblContacts.RowCount = 1;
            this.tblContacts.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tblContacts.Size = new System.Drawing.Size(1075, 106);
            this.tblContacts.TabIndex = 0;
            // 
            // ContactInfoRegion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.tblContacts);
            this.Name = "ContactInfoRegion";
            this.Size = new System.Drawing.Size(1075, 106);
            this.FormRegionShowing += new System.EventHandler(this.ContactInfoRegion_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.ContactInfoRegion_FormRegionClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "Quick Contact Info";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;
            manifest.ShowInspectorCompose = false;

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tblContacts;

        public partial class ContactInfoRegionFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public ContactInfoRegionFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                ContactInfoRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.ContactInfoRegionFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                ContactInfoRegion form = new ContactInfoRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal ContactInfoRegion ContactInfoRegion
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(ContactInfoRegion))
                        return (ContactInfoRegion)item;
                }
                return null;
            }
        }
    }
}
