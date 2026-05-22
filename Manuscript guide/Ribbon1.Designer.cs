namespace Manuscript_guide
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.ManuscriptGuide = this.Factory.CreateRibbonTab();
            this.groupDiagnostics = this.Factory.CreateRibbonGroup();
            this.btnPunc = this.Factory.CreateRibbonButton();
            this.btnCap = this.Factory.CreateRibbonButton();
            this.btnDash = this.Factory.CreateRibbonButton();
            this.btnData = this.Factory.CreateRibbonButton();
            this.btnItal = this.Factory.CreateRibbonButton();
            this.btnSub = this.Factory.CreateRibbonButton();
            this.btnWord = this.Factory.CreateRibbonButton();
            this.groupKeywordSearch = this.Factory.CreateRibbonGroup();
            this.btnKeyword = this.Factory.CreateRibbonButton();
            this.groupToolsSetup = this.Factory.CreateRibbonGroup();
            this.btnDetectAll = this.Factory.CreateRibbonButton();
            this.btnFixAll = this.Factory.CreateRibbonButton();
            this.btnClear = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.ManuscriptGuide.SuspendLayout();
            this.groupDiagnostics.SuspendLayout();
            this.groupKeywordSearch.SuspendLayout();
            this.groupToolsSetup.SuspendLayout();
            this.SuspendLayout();
            // 
            // ManuscriptGuide
            // 
            this.ManuscriptGuide.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ManuscriptGuide.Groups.Add(this.groupDiagnostics);
            this.ManuscriptGuide.Groups.Add(this.groupKeywordSearch);
            this.ManuscriptGuide.Groups.Add(this.groupToolsSetup);
            this.ManuscriptGuide.Label = "ManuGuide";
            this.ManuscriptGuide.Name = "ManuscriptGuide";
            this.ManuscriptGuide.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // groupDiagnostics
            // 
            this.groupDiagnostics.Items.Add(this.btnPunc);
            this.groupDiagnostics.Items.Add(this.btnCap);
            this.groupDiagnostics.Items.Add(this.btnDash);
            this.groupDiagnostics.Items.Add(this.btnData);
            this.groupDiagnostics.Items.Add(this.btnItal);
            this.groupDiagnostics.Items.Add(this.btnSub);
            this.groupDiagnostics.Items.Add(this.btnWord);
            this.groupDiagnostics.Label = "专项规范检测";
            this.groupDiagnostics.Name = "groupDiagnostics";
            // 
            // btnPunc
            // 
            this.btnPunc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPunc.Label = "标点与格式";
            this.btnPunc.Name = "btnPunc";
            this.btnPunc.OfficeImageId = "AutoCorrect";
            this.btnPunc.ScreenTip = "标点与格式";
            this.btnPunc.ShowImage = true;
            this.btnPunc.SuperTip = "检查中英文标点、半角空格、微符号编码和公式周边标点。";
            this.btnPunc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnCap
            // 
            this.btnCap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCap.Label = "大小写规范";
            this.btnCap.Name = "btnCap";
            this.btnCap.OfficeImageId = "ChangeCase";
            this.btnCap.ScreenTip = "大小写规范";
            this.btnCap.ShowImage = true;
            this.btnCap.SuperTip = "检查缩写首次定义、大小写一致性和图表交叉引用大写。";
            this.btnCap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnDash
            // 
            this.btnDash.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDash.Label = "横线与混排";
            this.btnDash.Name = "btnDash";
            this.btnDash.OfficeImageId = "ParagraphSpacingMenu";
            this.btnDash.ScreenTip = "横线与混排";
            this.btnDash.ShowImage = true;
            this.btnDash.SuperTip = "检查连字符、en dash、em dash、复合修饰语和混排空格。";
            this.btnDash.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnData
            // 
            this.btnData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnData.Label = "数据与数值";
            this.btnData.Name = "btnData";
            this.btnData.OfficeImageId = "EquationProfessional";
            this.btnData.ScreenTip = "数据与数值";
            this.btnData.ShowImage = true;
            this.btnData.SuperTip = "检查数值单位空格、科学计数法、范围表达和数学符号。";
            this.btnData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnItal
            // 
            this.btnItal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnItal.Label = "变量与斜体";
            this.btnItal.Name = "btnItal";
            this.btnItal.OfficeImageId = "Italic";
            this.btnItal.ScreenTip = "变量与斜体";
            this.btnItal.ShowImage = true;
            this.btnItal.SuperTip = "检查变量斜体、函数正体和不应斜体的学术短语。";
            this.btnItal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnSub
            // 
            this.btnSub.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSub.Label = "角标与改写";
            this.btnSub.Name = "btnSub";
            this.btnSub.OfficeImageId = "Subscript";
            this.btnSub.ScreenTip = "角标与改写";
            this.btnSub.ShowImage = true;
            this.btnSub.SuperTip = "检查化学式、单位幂次、描述性角标和 LaTeX 风格行内角标。";
            this.btnSub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // btnWord
            // 
            this.btnWord.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWord.Label = "语病与措辞";
            this.btnWord.Name = "btnWord";
            this.btnWord.OfficeImageId = "SpellingAndGrammar";
            this.btnWord.ScreenTip = "语病与措辞";
            this.btnWord.ShowImage = true;
            this.btnWord.SuperTip = "检查常见学术语病、措辞冗余和英文表达风险。";
            this.btnWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DiagnosticButton_Click);
            // 
            // groupKeywordSearch
            // 
            this.groupKeywordSearch.Items.Add(this.btnKeyword);
            this.groupKeywordSearch.Label = "快捷检索";
            this.groupKeywordSearch.Name = "groupKeywordSearch";
            // 
            // btnKeyword
            // 
            this.btnKeyword.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnKeyword.Label = "关键词标记";
            this.btnKeyword.Name = "btnKeyword";
            this.btnKeyword.OfficeImageId = "FindDialog";
            this.btnKeyword.ScreenTip = "关键词标记";
            this.btnKeyword.ShowImage = true;
            this.btnKeyword.SuperTip = "打开关键词检索面板，对指定术语进行全文高亮标记。";
            this.btnKeyword.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnKeyword_Click);
            // 
            // groupToolsSetup
            // 
            this.groupToolsSetup.Items.Add(this.btnDetectAll);
            this.groupToolsSetup.Items.Add(this.btnFixAll);
            this.groupToolsSetup.Items.Add(this.btnClear);
            this.groupToolsSetup.Items.Add(this.btnSettings);
            this.groupToolsSetup.Label = "一键与配置";
            this.groupToolsSetup.Name = "groupToolsSetup";
            // 
            // btnDetectAll
            // 
            this.btnDetectAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDetectAll.Label = "一键检测";
            this.btnDetectAll.Name = "btnDetectAll";
            this.btnDetectAll.OfficeImageId = "SpellingAndGrammar";
            this.btnDetectAll.ScreenTip = "一键检测";
            this.btnDetectAll.ShowImage = true;
            this.btnDetectAll.SuperTip = "连续运行全部专项检测，只生成问题队列和高亮标记，不直接修改正文。";
            this.btnDetectAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDetectAll_Click);
            // 
            // btnFixAll
            // 
            this.btnFixAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFixAll.Label = "一键修复";
            this.btnFixAll.Name = "btnFixAll";
            this.btnFixAll.OfficeImageId = "ReviewAcceptChange";
            this.btnFixAll.ScreenTip = "一键修复";
            this.btnFixAll.ShowImage = true;
            this.btnFixAll.SuperTip = "先运行全部专项检测，再批量应用可修复建议，并保留淡色底纹和 Word 批注。";
            this.btnFixAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFixAll_Click);
            // 
            // btnClear
            // 
            this.btnClear.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnClear.Label = "清除标记";
            this.btnClear.Name = "btnClear";
            this.btnClear.OfficeImageId = "ClearFormatting";
            this.btnClear.ScreenTip = "清除标记";
            this.btnClear.ShowImage = true;
            this.btnClear.SuperTip = "清除插件生成的诊断底纹和标记，保留或清除手动高亮取决于高级设置。";
            this.btnClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnClear_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Label = "高级设置";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.OfficeImageId = "FileOptions";
            this.btnSettings.ScreenTip = "高级设置";
            this.btnSettings.ShowImage = true;
            this.btnSettings.SuperTip = "打开规则、颜色、白名单、阈值和导入导出配置。";
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSettings_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ManuscriptGuide);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.ManuscriptGuide.ResumeLayout(false);
            this.ManuscriptGuide.PerformLayout();
            this.groupDiagnostics.ResumeLayout(false);
            this.groupDiagnostics.PerformLayout();
            this.groupKeywordSearch.ResumeLayout(false);
            this.groupKeywordSearch.PerformLayout();
            this.groupToolsSetup.ResumeLayout(false);
            this.groupToolsSetup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ManuscriptGuide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDiagnostics;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPunc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDash;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnItal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupKeywordSearch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnKeyword;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupToolsSetup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDetectAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFixAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
