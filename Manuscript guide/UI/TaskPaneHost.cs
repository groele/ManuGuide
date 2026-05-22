using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace Manuscript_guide.UI
{
    public class TaskPaneHost : UserControl
    {
        private ElementHost elementHost;
        private TaskPaneWPFControl wpfControl;

        public TaskPaneHost()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.elementHost = new ElementHost();
            this.wpfControl = new TaskPaneWPFControl();
            
            this.SuspendLayout();
            
            // 
            // elementHost
            // 
            this.elementHost.Dock = DockStyle.Fill;
            this.elementHost.Location = new System.Drawing.Point(0, 0);
            this.elementHost.Name = "elementHost";
            this.elementHost.Size = new System.Drawing.Size(480, 700);
            this.elementHost.TabIndex = 0;
            this.elementHost.Text = "elementHost";
            this.elementHost.Child = this.wpfControl;
            
            // 
            // TaskPaneHost
            // 
            this.Controls.Add(this.elementHost);
            this.Name = "TaskPaneHost";
            this.Size = new System.Drawing.Size(480, 700);
            
            this.ResumeLayout(false);
        }
    }
}
