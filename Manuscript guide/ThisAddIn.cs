using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Manuscript_guide
{
    public partial class ThisAddIn
    {
        private UI.TaskPaneHost taskPaneHost;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;

        public Microsoft.Office.Tools.CustomTaskPane CustomPane
        {
            get { return taskPane; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                taskPaneHost = new UI.TaskPaneHost();
                taskPane = this.CustomTaskPanes.Add(taskPaneHost, "学术论文排版审查 (Manuscript Guide)");
                taskPane.Width = 480;
                taskPane.Visible = true; // Auto-open on startup to showcase premium UI

                // Keep taskpane visible when scans or settings are requested
                Services.EventBus.ScanTriggered += (module) => { taskPane.Visible = true; };
                Services.EventBus.OpenSettingsRequested += () => { taskPane.Visible = true; };
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Manuscript Guide task pane initialization failed: " + ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
