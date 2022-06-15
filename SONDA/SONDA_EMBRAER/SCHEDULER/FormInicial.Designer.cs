namespace SCHEDULER
{
    partial class FormInicial
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.btnArchiveDiario = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBaterPonto = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAlterarSenhaSAP = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(30, 145);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Archive Diário";
            // 
            // btnArchiveDiario
            // 
            this.btnArchiveDiario.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnArchiveDiario.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnArchiveDiario.Location = new System.Drawing.Point(162, 137);
            this.btnArchiveDiario.Name = "btnArchiveDiario";
            this.btnArchiveDiario.Size = new System.Drawing.Size(79, 32);
            this.btnArchiveDiario.TabIndex = 1;
            this.btnArchiveDiario.Text = "Executar";
            this.btnArchiveDiario.UseVisualStyleBackColor = true;
            this.btnArchiveDiario.Click += new System.EventHandler(this.btnArchiveDiario_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(30, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "Bater Ponto";
            // 
            // btnBaterPonto
            // 
            this.btnBaterPonto.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBaterPonto.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnBaterPonto.Location = new System.Drawing.Point(162, 88);
            this.btnBaterPonto.Name = "btnBaterPonto";
            this.btnBaterPonto.Size = new System.Drawing.Size(79, 32);
            this.btnBaterPonto.TabIndex = 1;
            this.btnBaterPonto.Text = "Executar";
            this.btnBaterPonto.UseVisualStyleBackColor = true;
            this.btnBaterPonto.Click += new System.EventHandler(this.btnBaterPonto_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(217, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 25);
            this.label3.TabIndex = 2;
            this.label3.Text = "Scheduler";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(30, 194);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(111, 17);
            this.label4.TabIndex = 0;
            this.label4.Text = "Alterar senha SAP";
            // 
            // btnAlterarSenhaSAP
            // 
            this.btnAlterarSenhaSAP.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlterarSenhaSAP.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnAlterarSenhaSAP.Location = new System.Drawing.Point(162, 186);
            this.btnAlterarSenhaSAP.Name = "btnAlterarSenhaSAP";
            this.btnAlterarSenhaSAP.Size = new System.Drawing.Size(79, 32);
            this.btnAlterarSenhaSAP.TabIndex = 1;
            this.btnAlterarSenhaSAP.Text = "Executar";
            this.btnAlterarSenhaSAP.UseVisualStyleBackColor = true;
            this.btnAlterarSenhaSAP.Click += new System.EventHandler(this.btnAlterarSenhaSAP_Click);
            // 
            // FormInicial
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(530, 421);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnBaterPonto);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnAlterarSenhaSAP);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnArchiveDiario);
            this.Controls.Add(this.label1);
            this.Name = "FormInicial";
            this.Text = "Ferramentas SCHEDULER SONDA";
            this.Load += new System.EventHandler(this.FormInicial_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label label1;
        private Button btnArchiveDiario;
        private Label label2;
        private Button btnBaterPonto;
        private Label label3;
        private Label label4;
        private Button btnAlterarSenhaSAP;
    }
}