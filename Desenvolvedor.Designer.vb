<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Desenvolvedor
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Desenvolvedor))
        Me.cbbox_Excel = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.VoltarParaOMenuPrincipalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cbbox_Exportar = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tbp_POpBkb = New System.Windows.Forms.TabPage()
        Me.dgv_POpBkb = New System.Windows.Forms.DataGridView()
        Me.Processo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Início = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fim = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Total = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbp_POpFz = New System.Windows.Forms.TabPage()
        Me.dgv_POpFz = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbp_POpPlk = New System.Windows.Forms.TabPage()
        Me.dgv_POpPlk = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbp_LaborPlk = New System.Windows.Forms.TabPage()
        Me.dgv_PLbPlk = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbp_CmvPlk = New System.Windows.Forms.TabPage()
        Me.btn_Exportar = New System.Windows.Forms.Button()
        Me.dgv_PCmvPlk = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MenuStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.tbp_POpBkb.SuspendLayout()
        CType(Me.dgv_POpBkb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp_POpFz.SuspendLayout()
        CType(Me.dgv_POpFz, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp_POpPlk.SuspendLayout()
        CType(Me.dgv_POpPlk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp_LaborPlk.SuspendLayout()
        CType(Me.dgv_PLbPlk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbp_CmvPlk.SuspendLayout()
        CType(Me.dgv_PCmvPlk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbbox_Excel
        '
        Me.cbbox_Excel.AutoSize = True
        Me.cbbox_Excel.Font = New System.Drawing.Font("Century Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbbox_Excel.Location = New System.Drawing.Point(12, 51)
        Me.cbbox_Excel.Name = "cbbox_Excel"
        Me.cbbox_Excel.Size = New System.Drawing.Size(187, 25)
        Me.cbbox_Excel.TabIndex = 28
        Me.cbbox_Excel.Text = "Deixar Excel Visível"
        Me.cbbox_Excel.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VoltarParaOMenuPrincipalToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(767, 28)
        Me.MenuStrip1.TabIndex = 29
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'VoltarParaOMenuPrincipalToolStripMenuItem
        '
        Me.VoltarParaOMenuPrincipalToolStripMenuItem.Name = "VoltarParaOMenuPrincipalToolStripMenuItem"
        Me.VoltarParaOMenuPrincipalToolStripMenuItem.Size = New System.Drawing.Size(211, 24)
        Me.VoltarParaOMenuPrincipalToolStripMenuItem.Text = "Voltar para o Menu Principal"
        '
        'cbbox_Exportar
        '
        Me.cbbox_Exportar.AutoSize = True
        Me.cbbox_Exportar.Font = New System.Drawing.Font("Century Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbbox_Exportar.Location = New System.Drawing.Point(217, 51)
        Me.cbbox_Exportar.Name = "cbbox_Exportar"
        Me.cbbox_Exportar.Size = New System.Drawing.Size(255, 25)
        Me.cbbox_Exportar.TabIndex = 31
        Me.cbbox_Exportar.Text = "Exportar Report para Excel"
        Me.cbbox_Exportar.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tbp_POpBkb)
        Me.TabControl1.Controls.Add(Me.tbp_POpFz)
        Me.TabControl1.Controls.Add(Me.tbp_POpPlk)
        Me.TabControl1.Controls.Add(Me.tbp_LaborPlk)
        Me.TabControl1.Controls.Add(Me.tbp_CmvPlk)
        Me.TabControl1.Location = New System.Drawing.Point(12, 82)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(743, 477)
        Me.TabControl1.TabIndex = 32
        '
        'tbp_POpBkb
        '
        Me.tbp_POpBkb.Controls.Add(Me.dgv_POpBkb)
        Me.tbp_POpBkb.Location = New System.Drawing.Point(4, 25)
        Me.tbp_POpBkb.Name = "tbp_POpBkb"
        Me.tbp_POpBkb.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp_POpBkb.Size = New System.Drawing.Size(735, 448)
        Me.tbp_POpBkb.TabIndex = 0
        Me.tbp_POpBkb.Text = "Operacional BKB"
        Me.tbp_POpBkb.UseVisualStyleBackColor = True
        '
        'dgv_POpBkb
        '
        Me.dgv_POpBkb.AllowUserToAddRows = False
        Me.dgv_POpBkb.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.dgv_POpBkb.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgv_POpBkb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_POpBkb.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Processo, Me.Início, Me.Fim, Me.Total})
        Me.dgv_POpBkb.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_POpBkb.Location = New System.Drawing.Point(3, 3)
        Me.dgv_POpBkb.Name = "dgv_POpBkb"
        Me.dgv_POpBkb.ReadOnly = True
        Me.dgv_POpBkb.RowHeadersWidth = 51
        Me.dgv_POpBkb.RowTemplate.Height = 24
        Me.dgv_POpBkb.Size = New System.Drawing.Size(729, 442)
        Me.dgv_POpBkb.TabIndex = 0
        '
        'Processo
        '
        Me.Processo.HeaderText = "Processo"
        Me.Processo.MinimumWidth = 6
        Me.Processo.Name = "Processo"
        Me.Processo.ReadOnly = True
        Me.Processo.Width = 125
        '
        'Início
        '
        Me.Início.HeaderText = "Início"
        Me.Início.MinimumWidth = 6
        Me.Início.Name = "Início"
        Me.Início.ReadOnly = True
        Me.Início.Width = 125
        '
        'Fim
        '
        Me.Fim.HeaderText = "Fim"
        Me.Fim.MinimumWidth = 6
        Me.Fim.Name = "Fim"
        Me.Fim.ReadOnly = True
        Me.Fim.Width = 125
        '
        'Total
        '
        Me.Total.HeaderText = "Total"
        Me.Total.MinimumWidth = 6
        Me.Total.Name = "Total"
        Me.Total.ReadOnly = True
        Me.Total.Width = 125
        '
        'tbp_POpFz
        '
        Me.tbp_POpFz.Controls.Add(Me.dgv_POpFz)
        Me.tbp_POpFz.Location = New System.Drawing.Point(4, 25)
        Me.tbp_POpFz.Name = "tbp_POpFz"
        Me.tbp_POpFz.Padding = New System.Windows.Forms.Padding(3)
        Me.tbp_POpFz.Size = New System.Drawing.Size(735, 448)
        Me.tbp_POpFz.TabIndex = 1
        Me.tbp_POpFz.Text = "Operacional FZ"
        Me.tbp_POpFz.UseVisualStyleBackColor = True
        '
        'dgv_POpFz
        '
        Me.dgv_POpFz.AllowUserToAddRows = False
        Me.dgv_POpFz.AllowUserToDeleteRows = False
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.dgv_POpFz.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgv_POpFz.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_POpFz.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8})
        Me.dgv_POpFz.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_POpFz.Location = New System.Drawing.Point(3, 3)
        Me.dgv_POpFz.Name = "dgv_POpFz"
        Me.dgv_POpFz.ReadOnly = True
        Me.dgv_POpFz.RowHeadersWidth = 51
        Me.dgv_POpFz.RowTemplate.Height = 24
        Me.dgv_POpFz.Size = New System.Drawing.Size(729, 442)
        Me.dgv_POpFz.TabIndex = 1
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Processo"
        Me.DataGridViewTextBoxColumn5.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        Me.DataGridViewTextBoxColumn5.Width = 125
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Início"
        Me.DataGridViewTextBoxColumn6.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        Me.DataGridViewTextBoxColumn6.Width = 125
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Fim"
        Me.DataGridViewTextBoxColumn7.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        Me.DataGridViewTextBoxColumn7.Width = 125
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "Total"
        Me.DataGridViewTextBoxColumn8.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        Me.DataGridViewTextBoxColumn8.Width = 125
        '
        'tbp_POpPlk
        '
        Me.tbp_POpPlk.Controls.Add(Me.dgv_POpPlk)
        Me.tbp_POpPlk.Location = New System.Drawing.Point(4, 25)
        Me.tbp_POpPlk.Name = "tbp_POpPlk"
        Me.tbp_POpPlk.Size = New System.Drawing.Size(735, 448)
        Me.tbp_POpPlk.TabIndex = 2
        Me.tbp_POpPlk.Text = "Operacional PLK"
        Me.tbp_POpPlk.UseVisualStyleBackColor = True
        '
        'dgv_POpPlk
        '
        Me.dgv_POpPlk.AllowUserToAddRows = False
        Me.dgv_POpPlk.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.dgv_POpPlk.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgv_POpPlk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_POpPlk.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4})
        Me.dgv_POpPlk.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_POpPlk.Location = New System.Drawing.Point(0, 0)
        Me.dgv_POpPlk.Name = "dgv_POpPlk"
        Me.dgv_POpPlk.ReadOnly = True
        Me.dgv_POpPlk.RowHeadersWidth = 51
        Me.dgv_POpPlk.RowTemplate.Height = 24
        Me.dgv_POpPlk.Size = New System.Drawing.Size(735, 448)
        Me.dgv_POpPlk.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Processo"
        Me.DataGridViewTextBoxColumn1.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 125
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Início"
        Me.DataGridViewTextBoxColumn2.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 125
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Fim"
        Me.DataGridViewTextBoxColumn3.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        Me.DataGridViewTextBoxColumn3.Width = 125
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Total"
        Me.DataGridViewTextBoxColumn4.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.Width = 125
        '
        'tbp_LaborPlk
        '
        Me.tbp_LaborPlk.Controls.Add(Me.dgv_PLbPlk)
        Me.tbp_LaborPlk.Location = New System.Drawing.Point(4, 25)
        Me.tbp_LaborPlk.Name = "tbp_LaborPlk"
        Me.tbp_LaborPlk.Size = New System.Drawing.Size(735, 448)
        Me.tbp_LaborPlk.TabIndex = 4
        Me.tbp_LaborPlk.Text = "Labor PLK"
        Me.tbp_LaborPlk.UseVisualStyleBackColor = True
        '
        'dgv_PLbPlk
        '
        Me.dgv_PLbPlk.AllowUserToAddRows = False
        Me.dgv_PLbPlk.AllowUserToDeleteRows = False
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.dgv_PLbPlk.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgv_PLbPlk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_PLbPlk.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.DataGridViewTextBoxColumn11, Me.DataGridViewTextBoxColumn12})
        Me.dgv_PLbPlk.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_PLbPlk.Location = New System.Drawing.Point(0, 0)
        Me.dgv_PLbPlk.Name = "dgv_PLbPlk"
        Me.dgv_PLbPlk.ReadOnly = True
        Me.dgv_PLbPlk.RowHeadersWidth = 51
        Me.dgv_PLbPlk.RowTemplate.Height = 24
        Me.dgv_PLbPlk.Size = New System.Drawing.Size(735, 448)
        Me.dgv_PLbPlk.TabIndex = 2
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.HeaderText = "Processo"
        Me.DataGridViewTextBoxColumn9.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        Me.DataGridViewTextBoxColumn9.Width = 125
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.HeaderText = "Início"
        Me.DataGridViewTextBoxColumn10.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        Me.DataGridViewTextBoxColumn10.Width = 125
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.HeaderText = "Fim"
        Me.DataGridViewTextBoxColumn11.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.ReadOnly = True
        Me.DataGridViewTextBoxColumn11.Width = 125
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.HeaderText = "Total"
        Me.DataGridViewTextBoxColumn12.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.ReadOnly = True
        Me.DataGridViewTextBoxColumn12.Width = 125
        '
        'tbp_CmvPlk
        '
        Me.tbp_CmvPlk.Controls.Add(Me.dgv_PCmvPlk)
        Me.tbp_CmvPlk.Location = New System.Drawing.Point(4, 25)
        Me.tbp_CmvPlk.Name = "tbp_CmvPlk"
        Me.tbp_CmvPlk.Size = New System.Drawing.Size(735, 448)
        Me.tbp_CmvPlk.TabIndex = 3
        Me.tbp_CmvPlk.Text = "CMV PLK"
        Me.tbp_CmvPlk.UseVisualStyleBackColor = True
        '
        'btn_Exportar
        '
        Me.btn_Exportar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btn_Exportar.Location = New System.Drawing.Point(540, 46)
        Me.btn_Exportar.Name = "btn_Exportar"
        Me.btn_Exportar.Size = New System.Drawing.Size(208, 37)
        Me.btn_Exportar.TabIndex = 33
        Me.btn_Exportar.Text = "Exportar Excel"
        Me.btn_Exportar.UseVisualStyleBackColor = True
        '
        'dgv_PCmvPlk
        '
        Me.dgv_PCmvPlk.AllowUserToAddRows = False
        Me.dgv_PCmvPlk.AllowUserToDeleteRows = False
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.dgv_PCmvPlk.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle5
        Me.dgv_PCmvPlk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_PCmvPlk.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn13, Me.DataGridViewTextBoxColumn14, Me.DataGridViewTextBoxColumn15, Me.DataGridViewTextBoxColumn16})
        Me.dgv_PCmvPlk.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_PCmvPlk.Location = New System.Drawing.Point(0, 0)
        Me.dgv_PCmvPlk.Name = "dgv_PCmvPlk"
        Me.dgv_PCmvPlk.ReadOnly = True
        Me.dgv_PCmvPlk.RowHeadersWidth = 51
        Me.dgv_PCmvPlk.RowTemplate.Height = 24
        Me.dgv_PCmvPlk.Size = New System.Drawing.Size(735, 448)
        Me.dgv_PCmvPlk.TabIndex = 3
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.HeaderText = "Processo"
        Me.DataGridViewTextBoxColumn13.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        Me.DataGridViewTextBoxColumn13.Width = 125
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.HeaderText = "Início"
        Me.DataGridViewTextBoxColumn14.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        Me.DataGridViewTextBoxColumn14.ReadOnly = True
        Me.DataGridViewTextBoxColumn14.Width = 125
        '
        'DataGridViewTextBoxColumn15
        '
        Me.DataGridViewTextBoxColumn15.HeaderText = "Fim"
        Me.DataGridViewTextBoxColumn15.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn15.Name = "DataGridViewTextBoxColumn15"
        Me.DataGridViewTextBoxColumn15.ReadOnly = True
        Me.DataGridViewTextBoxColumn15.Width = 125
        '
        'DataGridViewTextBoxColumn16
        '
        Me.DataGridViewTextBoxColumn16.HeaderText = "Total"
        Me.DataGridViewTextBoxColumn16.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn16.Name = "DataGridViewTextBoxColumn16"
        Me.DataGridViewTextBoxColumn16.ReadOnly = True
        Me.DataGridViewTextBoxColumn16.Width = 125
        '
        'Desenvolvedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(767, 571)
        Me.Controls.Add(Me.btn_Exportar)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cbbox_Exportar)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.cbbox_Excel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Desenvolvedor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Extrair Horas"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.tbp_POpBkb.ResumeLayout(False)
        CType(Me.dgv_POpBkb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp_POpFz.ResumeLayout(False)
        CType(Me.dgv_POpFz, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp_POpPlk.ResumeLayout(False)
        CType(Me.dgv_POpPlk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp_LaborPlk.ResumeLayout(False)
        CType(Me.dgv_PLbPlk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbp_CmvPlk.ResumeLayout(False)
        CType(Me.dgv_PCmvPlk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cbbox_Excel As CheckBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents VoltarParaOMenuPrincipalToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents cbbox_Exportar As CheckBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents tbp_POpBkb As TabPage
    Friend WithEvents tbp_POpFz As TabPage
    Friend WithEvents tbp_POpPlk As TabPage
    Friend WithEvents tbp_CmvPlk As TabPage
    Friend WithEvents tbp_LaborPlk As TabPage
    Friend WithEvents dgv_POpBkb As DataGridView
    Friend WithEvents Processo As DataGridViewTextBoxColumn
    Friend WithEvents Início As DataGridViewTextBoxColumn
    Friend WithEvents Fim As DataGridViewTextBoxColumn
    Friend WithEvents Total As DataGridViewTextBoxColumn
    Friend WithEvents btn_Exportar As Button
    Friend WithEvents dgv_POpPlk As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As DataGridViewTextBoxColumn
    Friend WithEvents dgv_POpFz As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn5 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As DataGridViewTextBoxColumn
    Friend WithEvents dgv_PLbPlk As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn9 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As DataGridViewTextBoxColumn
    Friend WithEvents dgv_PCmvPlk As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn13 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn15 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn16 As DataGridViewTextBoxColumn
End Class
