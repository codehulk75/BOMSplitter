namespace BOMSplitter
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
            this.bomFileLabel = new System.Windows.Forms.Label();
            this.bomFileTextBox = new System.Windows.Forms.TextBox();
            this.openFileButton = new System.Windows.Forms.Button();
            this.bomGridView = new System.Windows.Forms.DataGridView();
            this.closeButton = new System.Windows.Forms.Button();
            this.splitFileLabel = new System.Windows.Forms.Label();
            this.splitFileTextBox = new System.Windows.Forms.TextBox();
            this.splitFileButton = new System.Windows.Forms.Button();
            this.doSplitsButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.bomGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // bomFileLabel
            // 
            this.bomFileLabel.AutoSize = true;
            this.bomFileLabel.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bomFileLabel.Location = new System.Drawing.Point(12, 13);
            this.bomFileLabel.Name = "bomFileLabel";
            this.bomFileLabel.Size = new System.Drawing.Size(79, 13);
            this.bomFileLabel.TabIndex = 0;
            this.bomFileLabel.Text = "BOM File";
            // 
            // bomFileTextBox
            // 
            this.bomFileTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bomFileTextBox.Location = new System.Drawing.Point(114, 9);
            this.bomFileTextBox.Name = "bomFileTextBox";
            this.bomFileTextBox.Size = new System.Drawing.Size(606, 20);
            this.bomFileTextBox.TabIndex = 1;
            // 
            // openFileButton
            // 
            this.openFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.openFileButton.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openFileButton.Location = new System.Drawing.Point(726, 9);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(53, 20);
            this.openFileButton.TabIndex = 2;
            this.openFileButton.Text = "...";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // bomGridView
            // 
            this.bomGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bomGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.bomGridView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.bomGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.bomGridView.Location = new System.Drawing.Point(15, 83);
            this.bomGridView.Name = "bomGridView";
            this.bomGridView.Size = new System.Drawing.Size(764, 351);
            this.bomGridView.TabIndex = 3;
            // 
            // closeButton
            // 
            this.closeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.closeButton.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closeButton.Location = new System.Drawing.Point(704, 440);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(75, 23);
            this.closeButton.TabIndex = 4;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // splitFileLabel
            // 
            this.splitFileLabel.AutoSize = true;
            this.splitFileLabel.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.splitFileLabel.Location = new System.Drawing.Point(12, 37);
            this.splitFileLabel.Name = "splitFileLabel";
            this.splitFileLabel.Size = new System.Drawing.Size(97, 13);
            this.splitFileLabel.TabIndex = 6;
            this.splitFileLabel.Text = "Split File";
            // 
            // splitFileTextBox
            // 
            this.splitFileTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitFileTextBox.Location = new System.Drawing.Point(114, 33);
            this.splitFileTextBox.Name = "splitFileTextBox";
            this.splitFileTextBox.Size = new System.Drawing.Size(606, 20);
            this.splitFileTextBox.TabIndex = 7;
            // 
            // splitFileButton
            // 
            this.splitFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.splitFileButton.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.splitFileButton.Location = new System.Drawing.Point(726, 33);
            this.splitFileButton.Name = "splitFileButton";
            this.splitFileButton.Size = new System.Drawing.Size(53, 20);
            this.splitFileButton.TabIndex = 8;
            this.splitFileButton.Text = "...";
            this.splitFileButton.UseVisualStyleBackColor = true;
            this.splitFileButton.Click += new System.EventHandler(this.splitFileButton_Click);
            // 
            // doSplitsButton
            // 
            this.doSplitsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.doSplitsButton.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.doSplitsButton.Location = new System.Drawing.Point(692, 54);
            this.doSplitsButton.Name = "doSplitsButton";
            this.doSplitsButton.Size = new System.Drawing.Size(87, 23);
            this.doSplitsButton.TabIndex = 9;
            this.doSplitsButton.Text = "Split BOM";
            this.doSplitsButton.UseVisualStyleBackColor = true;
            this.doSplitsButton.Click += new System.EventHandler(this.doSplitsButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(791, 471);
            this.Controls.Add(this.doSplitsButton);
            this.Controls.Add(this.splitFileButton);
            this.Controls.Add(this.splitFileTextBox);
            this.Controls.Add(this.splitFileLabel);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.bomGridView);
            this.Controls.Add(this.openFileButton);
            this.Controls.Add(this.bomFileTextBox);
            this.Controls.Add(this.bomFileLabel);
            this.Name = "Form1";
            this.Text = "BOM Splitter";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.bomGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label bomFileLabel;
        private System.Windows.Forms.TextBox bomFileTextBox;
        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.DataGridView bomGridView;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Label splitFileLabel;
        private System.Windows.Forms.TextBox splitFileTextBox;
        private System.Windows.Forms.Button splitFileButton;
        private System.Windows.Forms.Button doSplitsButton;
    }
}

