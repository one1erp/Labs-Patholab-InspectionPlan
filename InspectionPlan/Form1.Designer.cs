namespace InspectionPlan
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
            this.richSpellCtrl1 = new ONE1_richTextCtrl.RichSpellCtrl();
            this.SuspendLayout();
            // 
            // richSpellCtrl1
            // 
            this.richSpellCtrl1.DocumentRtl = System.Windows.Forms.RightToLeft.Yes;
            this.richSpellCtrl1.Location = new System.Drawing.Point(0, -1);
            this.richSpellCtrl1.Name = "richSpellCtrl1";
            this.richSpellCtrl1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.richSpellCtrl1.Size = new System.Drawing.Size(937, 578);
            this.richSpellCtrl1.TabIndex = 0;
            this.richSpellCtrl1.Load += new System.EventHandler(this.richSpellCtrl1_Load);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(937, 575);
            this.Controls.Add(this.richSpellCtrl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private ONE1_richTextCtrl.RichSpellCtrl richSpellCtrl1;
    }
}