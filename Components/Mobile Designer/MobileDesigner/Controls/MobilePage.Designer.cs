namespace MobileDesigner.Controls
{
    partial class MobilePage
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Main = new MobileDesigner.Controls.MainPanel();
            this.Header = new MobileDesigner.Controls.HeaderPanel();
            this.Footer = new MobileDesigner.Controls.FooterPanel();
            this.SuspendLayout();
            // 
            // Main
            // 
            this.Main.BackColor = System.Drawing.Color.White;
            this.Main.BackgroundImage = null;
            this.Main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Main.Location = new System.Drawing.Point(0, 57);
            this.Main.Name = "Main";
            this.Main.Size = new System.Drawing.Size(270, 268);
            this.Main.TabIndex = 1;
            // 
            // Header
            // 
            this.Header.BackColor = System.Drawing.Color.SteelBlue;
            this.Header.BackgroundImage = null;
            this.Header.Dock = System.Windows.Forms.DockStyle.Top;
            this.Header.Location = new System.Drawing.Point(0, 0);
            this.Header.LogoImage = null;
            this.Header.LogoImageOffset = new System.Drawing.Point(0, 0);
            this.Header.LogoImageSize = new System.Drawing.Size(0, 0);
            this.Header.Name = "Header";
            this.Header.Size = new System.Drawing.Size(270, 57);
            this.Header.TabIndex = 0;
            // 
            // Footer
            // 
            this.Footer.BackColor = System.Drawing.Color.SteelBlue;
            this.Footer.BackgroundImage = null;
            this.Footer.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Footer.Location = new System.Drawing.Point(0, 325);
            this.Footer.Name = "Footer";
            this.Footer.Size = new System.Drawing.Size(270, 57);
            this.Footer.TabIndex = 2;
            // 
            // MobilePage
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.Main);
            this.Controls.Add(this.Header);
            this.Controls.Add(this.Footer);
            this.Name = "MobilePage";
            this.Size = new System.Drawing.Size(270, 382);
            this.ResumeLayout(false);

        }

        #endregion

        public FooterPanel Footer;
        public HeaderPanel Header;
        public MainPanel Main;
    }
}
