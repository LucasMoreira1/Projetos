namespace Ferramentas_Scheduler
{
    partial class FormPrincipal
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblBaterPonto = new System.Windows.Forms.Label();
            this.lblArchiveDiario = new System.Windows.Forms.Label();
            this.lblAlterarSenhaSAP = new System.Windows.Forms.Label();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.btnBaterPonto = new System.Windows.Forms.Button();
            this.btnArchiveDiario = new System.Windows.Forms.Button();
            this.btnAlterarSenhaSAP = new System.Windows.Forms.Button();
            this.lblLogin = new System.Windows.Forms.Label();
            this.lblHorario = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblBaterPonto
            // 
            this.lblBaterPonto.AutoSize = true;
            this.lblBaterPonto.Location = new System.Drawing.Point(53, 87);
            this.lblBaterPonto.Name = "lblBaterPonto";
            this.lblBaterPonto.Size = new System.Drawing.Size(77, 16);
            this.lblBaterPonto.TabIndex = 0;
            this.lblBaterPonto.Text = "Bater Ponto";
            // 
            // lblArchiveDiario
            // 
            this.lblArchiveDiario.AutoSize = true;
            this.lblArchiveDiario.Location = new System.Drawing.Point(53, 129);
            this.lblArchiveDiario.Name = "lblArchiveDiario";
            this.lblArchiveDiario.Size = new System.Drawing.Size(91, 16);
            this.lblArchiveDiario.TabIndex = 1;
            this.lblArchiveDiario.Text = "Archive Diário";
            // 
            // lblAlterarSenhaSAP
            // 
            this.lblAlterarSenhaSAP.AutoSize = true;
            this.lblAlterarSenhaSAP.Location = new System.Drawing.Point(53, 171);
            this.lblAlterarSenhaSAP.Name = "lblAlterarSenhaSAP";
            this.lblAlterarSenhaSAP.Size = new System.Drawing.Size(118, 16);
            this.lblAlterarSenhaSAP.TabIndex = 1;
            this.lblAlterarSenhaSAP.Text = "Alterar Senha SAP";
            // 
            // lblTitulo
            // 
            this.lblTitulo.AutoSize = true;
            this.lblTitulo.Location = new System.Drawing.Point(149, 23);
            this.lblTitulo.Name = "lblTitulo";
            this.lblTitulo.Size = new System.Drawing.Size(90, 16);
            this.lblTitulo.TabIndex = 0;
            this.lblTitulo.Text = "SCHEDULER";
            // 
            // btnBaterPonto
            // 
            this.btnBaterPonto.Location = new System.Drawing.Point(260, 84);
            this.btnBaterPonto.Name = "btnBaterPonto";
            this.btnBaterPonto.Size = new System.Drawing.Size(75, 23);
            this.btnBaterPonto.TabIndex = 2;
            this.btnBaterPonto.Text = "Executar";
            this.btnBaterPonto.UseVisualStyleBackColor = true;
            this.btnBaterPonto.Click += new System.EventHandler(this.btnBaterPonto_Click);
            // 
            // btnArchiveDiario
            // 
            this.btnArchiveDiario.Location = new System.Drawing.Point(260, 126);
            this.btnArchiveDiario.Name = "btnArchiveDiario";
            this.btnArchiveDiario.Size = new System.Drawing.Size(75, 23);
            this.btnArchiveDiario.TabIndex = 2;
            this.btnArchiveDiario.Text = "Executar";
            this.btnArchiveDiario.UseVisualStyleBackColor = true;
            // 
            // btnAlterarSenhaSAP
            // 
            this.btnAlterarSenhaSAP.Location = new System.Drawing.Point(260, 168);
            this.btnAlterarSenhaSAP.Name = "btnAlterarSenhaSAP";
            this.btnAlterarSenhaSAP.Size = new System.Drawing.Size(75, 23);
            this.btnAlterarSenhaSAP.TabIndex = 2;
            this.btnAlterarSenhaSAP.Text = "Executar";
            this.btnAlterarSenhaSAP.UseVisualStyleBackColor = true;
            // 
            // lblLogin
            // 
            this.lblLogin.AutoSize = true;
            this.lblLogin.Location = new System.Drawing.Point(12, 9);
            this.lblLogin.Name = "lblLogin";
            this.lblLogin.Size = new System.Drawing.Size(40, 16);
            this.lblLogin.TabIndex = 3;
            this.lblLogin.Text = "Login";
            // 
            // lblHorario
            // 
            this.lblHorario.AutoSize = true;
            this.lblHorario.Location = new System.Drawing.Point(325, 9);
            this.lblHorario.Name = "lblHorario";
            this.lblHorario.Size = new System.Drawing.Size(52, 16);
            this.lblHorario.TabIndex = 3;
            this.lblHorario.Text = "Horário";
            // 
            // FormPrincipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(389, 275);
            this.Controls.Add(this.lblHorario);
            this.Controls.Add(this.lblLogin);
            this.Controls.Add(this.btnAlterarSenhaSAP);
            this.Controls.Add(this.btnArchiveDiario);
            this.Controls.Add(this.btnBaterPonto);
            this.Controls.Add(this.lblAlterarSenhaSAP);
            this.Controls.Add(this.lblArchiveDiario);
            this.Controls.Add(this.lblTitulo);
            this.Controls.Add(this.lblBaterPonto);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FormPrincipal";
            this.Text = "Ferramentas Scheduler";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblBaterPonto;
        private System.Windows.Forms.Label lblArchiveDiario;
        private System.Windows.Forms.Label lblAlterarSenhaSAP;
        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Button btnBaterPonto;
        private System.Windows.Forms.Button btnArchiveDiario;
        private System.Windows.Forms.Button btnAlterarSenhaSAP;
        private System.Windows.Forms.Label lblLogin;
        private System.Windows.Forms.Label lblHorario;
    }
}

