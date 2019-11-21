namespace MCDOutlookAddIn
{
    partial class MCDRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MCDRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.MCDTad = this.Factory.CreateRibbonTab();
            this.FirstGroup = this.Factory.CreateRibbonGroup();
            this.storeNumber = this.Factory.CreateRibbonButton();
            this.MCDTad.SuspendLayout();
            this.FirstGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // MCDTad
            // 
            this.MCDTad.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.MCDTad.Groups.Add(this.FirstGroup);
            this.MCDTad.Label = "McDonald\'s";
            this.MCDTad.Name = "MCDTad";
            // 
            // FirstGroup
            // 
            this.FirstGroup.Items.Add(this.storeNumber);
            this.FirstGroup.Label = "McDonald\'s";
            this.FirstGroup.Name = "FirstGroup";
            // 
            // storeNumber
            // 
            this.storeNumber.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.storeNumber.Image = global::MCDOutlookAddIn.Properties.Resources.MCDLogoForButtonS;
            this.storeNumber.Label = "Store Number";
            this.storeNumber.Name = "storeNumber";
            this.storeNumber.ShowImage = true;
            this.storeNumber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StoreNumber_Click);
            // 
            // MCDRibbon
            // 
            this.Name = "MCDRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" +
    "l.Read";
            this.Tabs.Add(this.MCDTad);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MCDRibbon_Load);
            this.MCDTad.ResumeLayout(false);
            this.MCDTad.PerformLayout();
            this.FirstGroup.ResumeLayout(false);
            this.FirstGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab MCDTad;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup FirstGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton storeNumber;
    }

    partial class ThisRibbonCollection
    {
        internal MCDRibbon MCDRibbon
        {
            get { return this.GetRibbon<MCDRibbon>(); }
        }
    }
}
