using MobileDesigner.Controls;

namespace MobileDesigner
{
    partial class DesignerForm
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
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DesignerForm));
            this.mainPanel = new System.Windows.Forms.Panel();
            this.designPanel = new System.Windows.Forms.Panel();
            this.leftPanel = new System.Windows.Forms.Panel();
            this.showForgotPasswordPage = new System.Windows.Forms.Button();
            this.showChangePasswordPage = new System.Windows.Forms.Button();
            this.showNewRegistrationPage = new System.Windows.Forms.Button();
            this.showTodoListPage = new System.Windows.Forms.Button();
            this.showHomePage = new System.Windows.Forms.Button();
            this.showLoginPage = new System.Windows.Forms.Button();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.propertyGridPanel = new System.Windows.Forms.Panel();
            this.propertyGrid = new System.Windows.Forms.PropertyGrid();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.saveToolButton = new System.Windows.Forms.ToolStripButton();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.focusControl = new MobileDesigner.Controls.SelectControl();
            this.mainPanel.SuspendLayout();
            this.designPanel.SuspendLayout();
            this.leftPanel.SuspendLayout();
            this.propertyGridPanel.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.BackColor = System.Drawing.Color.White;
            this.mainPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.mainPanel.Controls.Add(this.designPanel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(167, 31);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(441, 569);
            this.mainPanel.TabIndex = 2;
            // 
            // designPanel
            // 
            this.designPanel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.designPanel.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("designPanel.BackgroundImage")));
            this.designPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.designPanel.Controls.Add(this.focusControl);
            this.designPanel.Location = new System.Drawing.Point(25, 20);
            this.designPanel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.designPanel.Name = "designPanel";
            this.designPanel.Size = new System.Drawing.Size(388, 526);
            this.designPanel.TabIndex = 0;
            // 
            // leftPanel
            // 
            this.leftPanel.BackColor = System.Drawing.Color.White;
            this.leftPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.leftPanel.Controls.Add(this.showForgotPasswordPage);
            this.leftPanel.Controls.Add(this.showChangePasswordPage);
            this.leftPanel.Controls.Add(this.showNewRegistrationPage);
            this.leftPanel.Controls.Add(this.showTodoListPage);
            this.leftPanel.Controls.Add(this.showHomePage);
            this.leftPanel.Controls.Add(this.showLoginPage);
            this.leftPanel.Dock = System.Windows.Forms.DockStyle.Left;
            this.leftPanel.Location = new System.Drawing.Point(3, 31);
            this.leftPanel.Name = "leftPanel";
            this.leftPanel.Size = new System.Drawing.Size(160, 569);
            this.leftPanel.TabIndex = 1;
            // 
            // showForgotPasswordPage
            // 
            this.showForgotPasswordPage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showForgotPasswordPage.Location = new System.Drawing.Point(10, 165);
            this.showForgotPasswordPage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showForgotPasswordPage.Name = "showForgotPasswordPage";
            this.showForgotPasswordPage.Size = new System.Drawing.Size(137, 25);
            this.showForgotPasswordPage.TabIndex = 5;
            this.showForgotPasswordPage.Text = "&Forgot Username";
            this.showForgotPasswordPage.UseVisualStyleBackColor = true;
            // 
            // showChangePasswordPage
            // 
            this.showChangePasswordPage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showChangePasswordPage.Location = new System.Drawing.Point(10, 134);
            this.showChangePasswordPage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showChangePasswordPage.Name = "showChangePasswordPage";
            this.showChangePasswordPage.Size = new System.Drawing.Size(137, 25);
            this.showChangePasswordPage.TabIndex = 4;
            this.showChangePasswordPage.Text = "&Change Password";
            this.showChangePasswordPage.UseVisualStyleBackColor = true;
            // 
            // showNewRegistrationPage
            // 
            this.showNewRegistrationPage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showNewRegistrationPage.Location = new System.Drawing.Point(10, 103);
            this.showNewRegistrationPage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showNewRegistrationPage.Name = "showNewRegistrationPage";
            this.showNewRegistrationPage.Size = new System.Drawing.Size(137, 25);
            this.showNewRegistrationPage.TabIndex = 3;
            this.showNewRegistrationPage.Text = "&New Registration";
            this.showNewRegistrationPage.UseVisualStyleBackColor = true;
            // 
            // showTodoListPage
            // 
            this.showTodoListPage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showTodoListPage.Location = new System.Drawing.Point(10, 72);
            this.showTodoListPage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showTodoListPage.Name = "showTodoListPage";
            this.showTodoListPage.Size = new System.Drawing.Size(137, 25);
            this.showTodoListPage.TabIndex = 2;
            this.showTodoListPage.Text = "&To Do List";
            this.showTodoListPage.UseVisualStyleBackColor = true;
            // 
            // showHomePage
            // 
            this.showHomePage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showHomePage.Location = new System.Drawing.Point(10, 41);
            this.showHomePage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showHomePage.Name = "showHomePage";
            this.showHomePage.Size = new System.Drawing.Size(137, 25);
            this.showHomePage.TabIndex = 1;
            this.showHomePage.Text = "&Home";
            this.showHomePage.UseVisualStyleBackColor = true;
            // 
            // showLoginPage
            // 
            this.showLoginPage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.showLoginPage.Location = new System.Drawing.Point(10, 10);
            this.showLoginPage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.showLoginPage.Name = "showLoginPage";
            this.showLoginPage.Size = new System.Drawing.Size(137, 25);
            this.showLoginPage.TabIndex = 0;
            this.showLoginPage.Text = "&Login";
            this.showLoginPage.UseVisualStyleBackColor = true;
            // 
            // splitter1
            // 
            this.splitter1.Dock = System.Windows.Forms.DockStyle.Right;
            this.splitter1.Location = new System.Drawing.Point(608, 31);
            this.splitter1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(4, 569);
            this.splitter1.TabIndex = 3;
            this.splitter1.TabStop = false;
            // 
            // propertyGridPanel
            // 
            this.propertyGridPanel.Controls.Add(this.propertyGrid);
            this.propertyGridPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.propertyGridPanel.Location = new System.Drawing.Point(612, 31);
            this.propertyGridPanel.Name = "propertyGridPanel";
            this.propertyGridPanel.Size = new System.Drawing.Size(295, 569);
            this.propertyGridPanel.TabIndex = 2;
            // 
            // propertyGrid
            // 
            this.propertyGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertyGrid.HelpVisible = false;
            this.propertyGrid.Location = new System.Drawing.Point(0, 0);
            this.propertyGrid.Name = "propertyGrid";
            this.propertyGrid.Size = new System.Drawing.Size(295, 569);
            this.propertyGrid.TabIndex = 0;
            this.propertyGrid.ToolbarVisible = false;
            // 
            // toolStrip
            // 
            this.toolStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolButton});
            this.toolStrip.Location = new System.Drawing.Point(3, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.toolStrip.Size = new System.Drawing.Size(904, 31);
            this.toolStrip.TabIndex = 0;
            this.toolStrip.Text = "Standard Toolbar";
            // 
            // saveToolButton
            // 
            this.saveToolButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.saveToolButton.Image = ((System.Drawing.Image)(resources.GetObject("saveToolButton.Image")));
            this.saveToolButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolButton.Name = "saveToolButton";
            this.saveToolButton.Size = new System.Drawing.Size(28, 28);
            this.saveToolButton.Text = "Save";
            // 
            // splitter2
            // 
            this.splitter2.Enabled = false;
            this.splitter2.Location = new System.Drawing.Point(163, 31);
            this.splitter2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(4, 569);
            this.splitter2.TabIndex = 4;
            this.splitter2.TabStop = false;
            // 
            // focusControl
            // 
            this.focusControl.AllowClickThrough = false;
            this.focusControl.BackColor = System.Drawing.Color.Red;
            this.focusControl.Location = new System.Drawing.Point(35, 48);
            this.focusControl.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.focusControl.Name = "focusControl";
            this.focusControl.Size = new System.Drawing.Size(57, 52);
            this.focusControl.TabIndex = 0;
            this.focusControl.TabStop = false;
            this.focusControl.Visible = false;
            // 
            // DesignerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(910, 603);
            this.Controls.Add(this.mainPanel);
            this.Controls.Add(this.splitter2);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.leftPanel);
            this.Controls.Add(this.propertyGridPanel);
            this.Controls.Add(this.toolStrip);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DesignerForm";
            this.Padding = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mobile Designer";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DesignerFormFormClosing);
            this.Load += new System.EventHandler(this.DesignerFormLoad);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DesignerFormKeyDown);
            this.mainPanel.ResumeLayout(false);
            this.designPanel.ResumeLayout(false);
            this.leftPanel.ResumeLayout(false);
            this.propertyGridPanel.ResumeLayout(false);
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Panel leftPanel;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Button showForgotPasswordPage;
        private System.Windows.Forms.Button showChangePasswordPage;
        private System.Windows.Forms.Button showNewRegistrationPage;
        private System.Windows.Forms.Button showTodoListPage;
        private System.Windows.Forms.Button showHomePage;
        private System.Windows.Forms.Button showLoginPage;
        private System.Windows.Forms.Panel designPanel;
        private SelectControl focusControl;
        private System.Windows.Forms.Panel propertyGridPanel;
        private System.Windows.Forms.PropertyGrid propertyGrid;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ToolStripButton saveToolButton;
        private System.Windows.Forms.Splitter splitter2;
    }
}

