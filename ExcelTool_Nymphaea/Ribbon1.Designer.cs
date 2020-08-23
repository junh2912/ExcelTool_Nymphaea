namespace ExcelTool_Nymphaea
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_Nymphaea = this.Factory.CreateRibbonTab();
            this.group_InsertImage = this.Factory.CreateRibbonGroup();
            this.editBox_FilePath = this.Factory.CreateRibbonEditBox();
            this.editBox_FileNameCol = this.Factory.CreateRibbonEditBox();
            this.button_Insert = this.Factory.CreateRibbonButton();
            this.editBox_Extension = this.Factory.CreateRibbonEditBox();
            this.editBox_InsertCol = this.Factory.CreateRibbonEditBox();
            this.button_Delete = this.Factory.CreateRibbonButton();
            this.tab_Nymphaea.SuspendLayout();
            this.group_InsertImage.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Nymphaea
            // 
            this.tab_Nymphaea.Groups.Add(this.group_InsertImage);
            this.tab_Nymphaea.Label = "Nymphaea Tool";
            this.tab_Nymphaea.Name = "tab_Nymphaea";
            // 
            // group_InsertImage
            // 
            this.group_InsertImage.Items.Add(this.editBox_FilePath);
            this.group_InsertImage.Items.Add(this.editBox_FileNameCol);
            this.group_InsertImage.Items.Add(this.button_Insert);
            this.group_InsertImage.Items.Add(this.editBox_Extension);
            this.group_InsertImage.Items.Add(this.editBox_InsertCol);
            this.group_InsertImage.Items.Add(this.button_Delete);
            this.group_InsertImage.Label = "Insert Image";
            this.group_InsertImage.Name = "group_InsertImage";
            // 
            // editBox_FilePath
            // 
            this.editBox_FilePath.Label = "File Path";
            this.editBox_FilePath.Name = "editBox_FilePath";
            this.editBox_FilePath.Text = null;
            // 
            // editBox_FileNameCol
            // 
            this.editBox_FileNameCol.Label = "File Name Col";
            this.editBox_FileNameCol.Name = "editBox_FileNameCol";
            this.editBox_FileNameCol.Text = null;
            // 
            // button_Insert
            // 
            this.button_Insert.Label = "Insert";
            this.button_Insert.Name = "button_Insert";
            this.button_Insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Insert_Click);
            // 
            // editBox_Extension
            // 
            this.editBox_Extension.Label = "Extension";
            this.editBox_Extension.Name = "editBox_Extension";
            this.editBox_Extension.Text = null;
            // 
            // editBox_InsertCol
            // 
            this.editBox_InsertCol.Label = "Insert Col";
            this.editBox_InsertCol.Name = "editBox_InsertCol";
            this.editBox_InsertCol.Text = null;
            // 
            // button_Delete
            // 
            this.button_Delete.Label = "Delete";
            this.button_Delete.Name = "button_Delete";
            this.button_Delete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Delete_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_Nymphaea);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_Nymphaea.ResumeLayout(false);
            this.tab_Nymphaea.PerformLayout();
            this.group_InsertImage.ResumeLayout(false);
            this.group_InsertImage.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_Nymphaea;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_InsertImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_FilePath;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_FileNameCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Insert;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_Extension;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_InsertCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Delete;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
