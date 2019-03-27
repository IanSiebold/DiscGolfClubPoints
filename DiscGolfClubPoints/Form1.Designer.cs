namespace DiscGolfClubPoints
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.enterButton = new System.Windows.Forms.Button();
            this.nameCombo = new System.Windows.Forms.ComboBox();
            this.updateButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(58, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter your first and last name";
            // 
            // enterButton
            // 
            this.enterButton.Location = new System.Drawing.Point(397, 59);
            this.enterButton.Name = "enterButton";
            this.enterButton.Size = new System.Drawing.Size(75, 23);
            this.enterButton.TabIndex = 2;
            this.enterButton.Text = "Enter";
            this.enterButton.UseVisualStyleBackColor = true;
            this.enterButton.Click += new System.EventHandler(this.enterButton_Click);
            // 
            // nameCombo
            // 
            this.nameCombo.FormattingEnabled = true;
            this.nameCombo.Location = new System.Drawing.Point(61, 59);
            this.nameCombo.Name = "nameCombo";
            this.nameCombo.Size = new System.Drawing.Size(330, 21);
            this.nameCombo.TabIndex = 3;
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(450, 12);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 4;
            this.updateButton.Text = "Update";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 124);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(this.nameCombo);
            this.Controls.Add(this.enterButton);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Disc Golf Club";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button enterButton;
        private System.Windows.Forms.ComboBox nameCombo;
        private System.Windows.Forms.Button updateButton;
    }
}

