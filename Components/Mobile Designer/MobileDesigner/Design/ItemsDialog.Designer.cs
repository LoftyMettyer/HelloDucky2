namespace MobileDesigner.Design
{
    partial class ItemsDialog
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
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.availableItemsList = new System.Windows.Forms.ListBox();
            this.selectedItemsList = new System.Windows.Forms.ListBox();
            this.addButton = new System.Windows.Forms.Button();
            this.removeButton = new System.Windows.Forms.Button();
            this.moveUpButton = new System.Windows.Forms.Button();
            this.moveDownButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(622, 314);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 24);
            this.buttonCancel.TabIndex = 9;
            this.buttonCancel.Text = "&Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(541, 314);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 24);
            this.buttonOK.TabIndex = 8;
            this.buttonOK.Text = "&OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            // 
            // availableItemsList
            // 
            this.availableItemsList.FormattingEnabled = true;
            this.availableItemsList.Location = new System.Drawing.Point(13, 27);
            this.availableItemsList.Name = "availableItemsList";
            this.availableItemsList.Size = new System.Drawing.Size(263, 277);
            this.availableItemsList.Sorted = true;
            this.availableItemsList.TabIndex = 1;
            this.availableItemsList.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.AvailableItemsListMouseDoubleClick);
            // 
            // selectedItemsList
            // 
            this.selectedItemsList.FormattingEnabled = true;
            this.selectedItemsList.Location = new System.Drawing.Point(366, 27);
            this.selectedItemsList.Name = "selectedItemsList";
            this.selectedItemsList.Size = new System.Drawing.Size(263, 277);
            this.selectedItemsList.TabIndex = 5;
            this.selectedItemsList.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.SelectedItemsListMouseDoubleClick);
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(283, 27);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(76, 24);
            this.addButton.TabIndex = 2;
            this.addButton.Text = "&Add";
            this.addButton.UseVisualStyleBackColor = true;
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(283, 57);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(76, 24);
            this.removeButton.TabIndex = 3;
            this.removeButton.Text = "&Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            // 
            // moveUpButton
            // 
            this.moveUpButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.moveUpButton.Location = new System.Drawing.Point(638, 27);
            this.moveUpButton.Name = "moveUpButton";
            this.moveUpButton.Size = new System.Drawing.Size(59, 24);
            this.moveUpButton.TabIndex = 6;
            this.moveUpButton.Text = "&Up";
            this.moveUpButton.UseVisualStyleBackColor = true;
            // 
            // moveDownButton
            // 
            this.moveDownButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.moveDownButton.Location = new System.Drawing.Point(638, 57);
            this.moveDownButton.Name = "moveDownButton";
            this.moveDownButton.Size = new System.Drawing.Size(59, 24);
            this.moveDownButton.TabIndex = 7;
            this.moveDownButton.Text = "&Down";
            this.moveDownButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Available Items:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(363, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Selected Items:";
            // 
            // ItemsDialog
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(709, 347);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.moveDownButton);
            this.Controls.Add(this.moveUpButton);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.selectedItemsList);
            this.Controls.Add(this.availableItemsList);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ItemsDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Items";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ItemsDialogFormClosing);
            this.Load += new System.EventHandler(this.ItemsDialogLoad);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.ListBox availableItemsList;
        private System.Windows.Forms.ListBox selectedItemsList;
        private System.Windows.Forms.Button addButton;
        private System.Windows.Forms.Button removeButton;
        private System.Windows.Forms.Button moveUpButton;
        private System.Windows.Forms.Button moveDownButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}