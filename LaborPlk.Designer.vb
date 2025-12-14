<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_LaborPlk
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(form_LaborPlk))
        Me.lbNome_BHoras = New System.Windows.Forms.Label()
        Me.btn_ArqBHoras = New System.Windows.Forms.Button()
        Me.lbNome_Rel11Hrs = New System.Windows.Forms.Label()
        Me.lbNome_PtDia = New System.Windows.Forms.Label()
        Me.lbNome_EstMarc = New System.Windows.Forms.Label()
        Me.lbNome_CadFunc = New System.Windows.Forms.Label()
        Me.lbNome_Rel2Hrs = New System.Windows.Forms.Label()
        Me.btn_ArqRel2Hrs = New System.Windows.Forms.Button()
        Me.btn_ArqCadFunc = New System.Windows.Forms.Button()
        Me.btn_ArqRel11Hrs = New System.Windows.Forms.Button()
        Me.btn_ArqPtDia = New System.Windows.Forms.Button()
        Me.btn_ArqEstMarc = New System.Windows.Forms.Button()
        Me.btn_GerarPainel = New System.Windows.Forms.Button()
        Me.lbNome_PLbPlk = New System.Windows.Forms.Label()
        Me.btn_ArqPLbPlk = New System.Windows.Forms.Button()
        Me.lbId_EstMarc = New System.Windows.Forms.Label()
        Me.lb_EstMarc = New System.Windows.Forms.Label()
        Me.lb_Rel2Hrs = New System.Windows.Forms.Label()
        Me.lbId_Rel2Hrs = New System.Windows.Forms.Label()
        Me.lbId_CadFunc = New System.Windows.Forms.Label()
        Me.lbId_PtDia = New System.Windows.Forms.Label()
        Me.lbId_Rel11Hrs = New System.Windows.Forms.Label()
        Me.lbId_PLbPlk = New System.Windows.Forms.Label()
        Me.lb_CadFunc = New System.Windows.Forms.Label()
        Me.lb_PtDia = New System.Windows.Forms.Label()
        Me.lb_Rel11Hrs = New System.Windows.Forms.Label()
        Me.lb_PLbPlk = New System.Windows.Forms.Label()
        Me.lbId_BHoras = New System.Windows.Forms.Label()
        Me.lb_BHoras = New System.Windows.Forms.Label()
        Me.ofd1 = New System.Windows.Forms.OpenFileDialog()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btn_Minimizar = New System.Windows.Forms.Button()
        Me.btn_Fechar = New System.Windows.Forms.Button()
        Me.btn_Home = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbNome_BHoras
        '
        Me.lbNome_BHoras.AutoSize = True
        Me.lbNome_BHoras.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_BHoras.Location = New System.Drawing.Point(493, 502)
        Me.lbNome_BHoras.Name = "lbNome_BHoras"
        Me.lbNome_BHoras.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_BHoras.TabIndex = 20
        '
        'btn_ArqBHoras
        '
        Me.btn_ArqBHoras.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqBHoras.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqBHoras.FlatAppearance.BorderSize = 0
        Me.btn_ArqBHoras.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqBHoras.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqBHoras.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqBHoras.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqBHoras.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqBHoras.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqBHoras.Location = New System.Drawing.Point(486, 434)
        Me.btn_ArqBHoras.Name = "btn_ArqBHoras"
        Me.btn_ArqBHoras.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqBHoras.TabIndex = 19
        Me.btn_ArqBHoras.Text = "Resumo Banco de Horas"
        Me.btn_ArqBHoras.UseVisualStyleBackColor = False
        '
        'lbNome_Rel11Hrs
        '
        Me.lbNome_Rel11Hrs.AutoSize = True
        Me.lbNome_Rel11Hrs.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_Rel11Hrs.Location = New System.Drawing.Point(493, 411)
        Me.lbNome_Rel11Hrs.Name = "lbNome_Rel11Hrs"
        Me.lbNome_Rel11Hrs.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_Rel11Hrs.TabIndex = 18
        '
        'lbNome_PtDia
        '
        Me.lbNome_PtDia.AutoSize = True
        Me.lbNome_PtDia.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_PtDia.Location = New System.Drawing.Point(34, 595)
        Me.lbNome_PtDia.Name = "lbNome_PtDia"
        Me.lbNome_PtDia.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_PtDia.TabIndex = 17
        '
        'lbNome_EstMarc
        '
        Me.lbNome_EstMarc.AutoSize = True
        Me.lbNome_EstMarc.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_EstMarc.Location = New System.Drawing.Point(34, 502)
        Me.lbNome_EstMarc.Name = "lbNome_EstMarc"
        Me.lbNome_EstMarc.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_EstMarc.TabIndex = 16
        '
        'lbNome_CadFunc
        '
        Me.lbNome_CadFunc.AutoSize = True
        Me.lbNome_CadFunc.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_CadFunc.Location = New System.Drawing.Point(34, 411)
        Me.lbNome_CadFunc.Name = "lbNome_CadFunc"
        Me.lbNome_CadFunc.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_CadFunc.TabIndex = 15
        '
        'lbNome_Rel2Hrs
        '
        Me.lbNome_Rel2Hrs.AutoSize = True
        Me.lbNome_Rel2Hrs.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_Rel2Hrs.Location = New System.Drawing.Point(493, 318)
        Me.lbNome_Rel2Hrs.Name = "lbNome_Rel2Hrs"
        Me.lbNome_Rel2Hrs.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_Rel2Hrs.TabIndex = 13
        '
        'btn_ArqRel2Hrs
        '
        Me.btn_ArqRel2Hrs.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel2Hrs.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqRel2Hrs.FlatAppearance.BorderSize = 0
        Me.btn_ArqRel2Hrs.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel2Hrs.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel2Hrs.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel2Hrs.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqRel2Hrs.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqRel2Hrs.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqRel2Hrs.Location = New System.Drawing.Point(486, 248)
        Me.btn_ArqRel2Hrs.Name = "btn_ArqRel2Hrs"
        Me.btn_ArqRel2Hrs.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqRel2Hrs.TabIndex = 12
        Me.btn_ArqRel2Hrs.Text = "Mais de 2 Horas"
        Me.btn_ArqRel2Hrs.UseVisualStyleBackColor = False
        '
        'btn_ArqCadFunc
        '
        Me.btn_ArqCadFunc.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqCadFunc.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqCadFunc.FlatAppearance.BorderSize = 0
        Me.btn_ArqCadFunc.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqCadFunc.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqCadFunc.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqCadFunc.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqCadFunc.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqCadFunc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqCadFunc.Location = New System.Drawing.Point(28, 339)
        Me.btn_ArqCadFunc.Name = "btn_ArqCadFunc"
        Me.btn_ArqCadFunc.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqCadFunc.TabIndex = 9
        Me.btn_ArqCadFunc.Text = "Cadastro de Funcionários"
        Me.btn_ArqCadFunc.UseVisualStyleBackColor = False
        '
        'btn_ArqRel11Hrs
        '
        Me.btn_ArqRel11Hrs.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel11Hrs.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqRel11Hrs.FlatAppearance.BorderSize = 0
        Me.btn_ArqRel11Hrs.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel11Hrs.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel11Hrs.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqRel11Hrs.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqRel11Hrs.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqRel11Hrs.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqRel11Hrs.Location = New System.Drawing.Point(486, 339)
        Me.btn_ArqRel11Hrs.Name = "btn_ArqRel11Hrs"
        Me.btn_ArqRel11Hrs.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqRel11Hrs.TabIndex = 8
        Me.btn_ArqRel11Hrs.Text = "Relatório 11 Horas"
        Me.btn_ArqRel11Hrs.UseVisualStyleBackColor = False
        '
        'btn_ArqPtDia
        '
        Me.btn_ArqPtDia.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqPtDia.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqPtDia.FlatAppearance.BorderSize = 0
        Me.btn_ArqPtDia.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqPtDia.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqPtDia.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqPtDia.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqPtDia.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqPtDia.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqPtDia.Location = New System.Drawing.Point(28, 525)
        Me.btn_ArqPtDia.Name = "btn_ArqPtDia"
        Me.btn_ArqPtDia.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqPtDia.TabIndex = 7
        Me.btn_ArqPtDia.Text = "Ponto Diário"
        Me.btn_ArqPtDia.UseVisualStyleBackColor = False
        '
        'btn_ArqEstMarc
        '
        Me.btn_ArqEstMarc.BackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqEstMarc.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqEstMarc.FlatAppearance.BorderSize = 0
        Me.btn_ArqEstMarc.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqEstMarc.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqEstMarc.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(69, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqEstMarc.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqEstMarc.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqEstMarc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqEstMarc.Location = New System.Drawing.Point(28, 434)
        Me.btn_ArqEstMarc.Name = "btn_ArqEstMarc"
        Me.btn_ArqEstMarc.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqEstMarc.TabIndex = 3
        Me.btn_ArqEstMarc.Text = "Estatística das Marcações"
        Me.btn_ArqEstMarc.UseVisualStyleBackColor = False
        '
        'btn_GerarPainel
        '
        Me.btn_GerarPainel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_GerarPainel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_GerarPainel.FlatAppearance.BorderSize = 0
        Me.btn_GerarPainel.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_GerarPainel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_GerarPainel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(101, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_GerarPainel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_GerarPainel.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_GerarPainel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_GerarPainel.Location = New System.Drawing.Point(28, 618)
        Me.btn_GerarPainel.Name = "btn_GerarPainel"
        Me.btn_GerarPainel.Size = New System.Drawing.Size(905, 65)
        Me.btn_GerarPainel.TabIndex = 45
        Me.btn_GerarPainel.Text = "Gerar Painel"
        Me.btn_GerarPainel.UseVisualStyleBackColor = False
        '
        'lbNome_PLbPlk
        '
        Me.lbNome_PLbPlk.AutoSize = True
        Me.lbNome_PLbPlk.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNome_PLbPlk.Location = New System.Drawing.Point(34, 316)
        Me.lbNome_PLbPlk.Name = "lbNome_PLbPlk"
        Me.lbNome_PLbPlk.Size = New System.Drawing.Size(0, 20)
        Me.lbNome_PLbPlk.TabIndex = 14
        '
        'btn_ArqPLbPlk
        '
        Me.btn_ArqPLbPlk.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_ArqPLbPlk.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_ArqPLbPlk.FlatAppearance.BorderSize = 0
        Me.btn_ArqPLbPlk.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_ArqPLbPlk.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.btn_ArqPLbPlk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(101, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_ArqPLbPlk.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_ArqPLbPlk.Font = New System.Drawing.Font("Chicken Sans", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ArqPLbPlk.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_ArqPLbPlk.Location = New System.Drawing.Point(28, 248)
        Me.btn_ArqPLbPlk.Name = "btn_ArqPLbPlk"
        Me.btn_ArqPLbPlk.Size = New System.Drawing.Size(445, 65)
        Me.btn_ArqPLbPlk.TabIndex = 0
        Me.btn_ArqPLbPlk.Text = "Painel Labor PLK"
        Me.btn_ArqPLbPlk.UseVisualStyleBackColor = False
        '
        'lbId_EstMarc
        '
        Me.lbId_EstMarc.AutoSize = True
        Me.lbId_EstMarc.Location = New System.Drawing.Point(1375, 642)
        Me.lbId_EstMarc.Name = "lbId_EstMarc"
        Me.lbId_EstMarc.Size = New System.Drawing.Size(89, 17)
        Me.lbId_EstMarc.TabIndex = 90
        Me.lbId_EstMarc.Text = "lbId_EstMarc"
        '
        'lb_EstMarc
        '
        Me.lb_EstMarc.AutoSize = True
        Me.lb_EstMarc.Location = New System.Drawing.Point(1283, 642)
        Me.lb_EstMarc.Name = "lb_EstMarc"
        Me.lb_EstMarc.Size = New System.Drawing.Size(78, 17)
        Me.lb_EstMarc.TabIndex = 89
        Me.lb_EstMarc.Text = "lb_EstMarc"
        Me.lb_EstMarc.Visible = False
        '
        'lb_Rel2Hrs
        '
        Me.lb_Rel2Hrs.AutoSize = True
        Me.lb_Rel2Hrs.Location = New System.Drawing.Point(1283, 608)
        Me.lb_Rel2Hrs.Name = "lb_Rel2Hrs"
        Me.lb_Rel2Hrs.Size = New System.Drawing.Size(78, 17)
        Me.lb_Rel2Hrs.TabIndex = 88
        Me.lb_Rel2Hrs.Text = "lb_Rel2Hrs"
        Me.lb_Rel2Hrs.Visible = False
        '
        'lbId_Rel2Hrs
        '
        Me.lbId_Rel2Hrs.AutoSize = True
        Me.lbId_Rel2Hrs.Location = New System.Drawing.Point(1375, 608)
        Me.lbId_Rel2Hrs.Name = "lbId_Rel2Hrs"
        Me.lbId_Rel2Hrs.Size = New System.Drawing.Size(89, 17)
        Me.lbId_Rel2Hrs.TabIndex = 87
        Me.lbId_Rel2Hrs.Text = "lbId_Rel2Hrs"
        '
        'lbId_CadFunc
        '
        Me.lbId_CadFunc.AutoSize = True
        Me.lbId_CadFunc.Location = New System.Drawing.Point(1375, 625)
        Me.lbId_CadFunc.Name = "lbId_CadFunc"
        Me.lbId_CadFunc.Size = New System.Drawing.Size(94, 17)
        Me.lbId_CadFunc.TabIndex = 86
        Me.lbId_CadFunc.Text = "lbId_CadFunc"
        '
        'lbId_PtDia
        '
        Me.lbId_PtDia.AutoSize = True
        Me.lbId_PtDia.Location = New System.Drawing.Point(1375, 659)
        Me.lbId_PtDia.Name = "lbId_PtDia"
        Me.lbId_PtDia.Size = New System.Drawing.Size(72, 17)
        Me.lbId_PtDia.TabIndex = 85
        Me.lbId_PtDia.Text = "lbId_PtDia"
        '
        'lbId_Rel11Hrs
        '
        Me.lbId_Rel11Hrs.AutoSize = True
        Me.lbId_Rel11Hrs.Location = New System.Drawing.Point(1375, 676)
        Me.lbId_Rel11Hrs.Name = "lbId_Rel11Hrs"
        Me.lbId_Rel11Hrs.Size = New System.Drawing.Size(97, 17)
        Me.lbId_Rel11Hrs.TabIndex = 84
        Me.lbId_Rel11Hrs.Text = "lbId_Rel11Hrs"
        '
        'lbId_PLbPlk
        '
        Me.lbId_PLbPlk.AutoSize = True
        Me.lbId_PLbPlk.Location = New System.Drawing.Point(1375, 591)
        Me.lbId_PLbPlk.Name = "lbId_PLbPlk"
        Me.lbId_PLbPlk.Size = New System.Drawing.Size(82, 17)
        Me.lbId_PLbPlk.TabIndex = 83
        Me.lbId_PLbPlk.Text = "lbId_PLbPlk"
        '
        'lb_CadFunc
        '
        Me.lb_CadFunc.AutoSize = True
        Me.lb_CadFunc.Location = New System.Drawing.Point(1277, 625)
        Me.lb_CadFunc.Name = "lb_CadFunc"
        Me.lb_CadFunc.Size = New System.Drawing.Size(83, 17)
        Me.lb_CadFunc.TabIndex = 82
        Me.lb_CadFunc.Text = "lb_CadFunc"
        Me.lb_CadFunc.Visible = False
        '
        'lb_PtDia
        '
        Me.lb_PtDia.AutoSize = True
        Me.lb_PtDia.Location = New System.Drawing.Point(1300, 659)
        Me.lb_PtDia.Name = "lb_PtDia"
        Me.lb_PtDia.Size = New System.Drawing.Size(61, 17)
        Me.lb_PtDia.TabIndex = 81
        Me.lb_PtDia.Text = "lb_PtDia"
        Me.lb_PtDia.Visible = False
        '
        'lb_Rel11Hrs
        '
        Me.lb_Rel11Hrs.AutoSize = True
        Me.lb_Rel11Hrs.Location = New System.Drawing.Point(1274, 676)
        Me.lb_Rel11Hrs.Name = "lb_Rel11Hrs"
        Me.lb_Rel11Hrs.Size = New System.Drawing.Size(86, 17)
        Me.lb_Rel11Hrs.TabIndex = 80
        Me.lb_Rel11Hrs.Text = "lb_Rel11Hrs"
        Me.lb_Rel11Hrs.Visible = False
        '
        'lb_PLbPlk
        '
        Me.lb_PLbPlk.AutoSize = True
        Me.lb_PLbPlk.Location = New System.Drawing.Point(1290, 591)
        Me.lb_PLbPlk.Name = "lb_PLbPlk"
        Me.lb_PLbPlk.Size = New System.Drawing.Size(71, 17)
        Me.lb_PLbPlk.TabIndex = 79
        Me.lb_PLbPlk.Text = "lb_PLbPlk"
        Me.lb_PLbPlk.Visible = False
        '
        'lbId_BHoras
        '
        Me.lbId_BHoras.AutoSize = True
        Me.lbId_BHoras.Location = New System.Drawing.Point(1375, 693)
        Me.lbId_BHoras.Name = "lbId_BHoras"
        Me.lbId_BHoras.Size = New System.Drawing.Size(85, 17)
        Me.lbId_BHoras.TabIndex = 92
        Me.lbId_BHoras.Text = "lbId_BHoras"
        '
        'lb_BHoras
        '
        Me.lb_BHoras.AutoSize = True
        Me.lb_BHoras.Location = New System.Drawing.Point(1286, 693)
        Me.lb_BHoras.Name = "lb_BHoras"
        Me.lb_BHoras.Size = New System.Drawing.Size(74, 17)
        Me.lb_BHoras.TabIndex = 91
        Me.lb_BHoras.Text = "lb_BHoras"
        Me.lb_BHoras.Visible = False
        '
        'ofd1
        '
        Me.ofd1.Filter = "Planilhas do Excel|*xls;*xlsx;*xlsm;*csv;*all"
        Me.ofd1.InitialDirectory = "C:\Área de Trabalho"
        Me.ofd1.RestoreDirectory = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.Rotinas_BK_Brasil.My.Resources.Resources.PLK_Branco
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Location = New System.Drawing.Point(345, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(281, 115)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.Panel1.Controls.Add(Me.btn_Home)
        Me.Panel1.Controls.Add(Me.btn_Minimizar)
        Me.Panel1.Controls.Add(Me.btn_Fechar)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(960, 120)
        Me.Panel1.TabIndex = 93
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Chicken Sans", 32.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(242, Byte), Integer), CType(CType(86, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(228, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(505, 65)
        Me.Label2.TabIndex = 94
        Me.Label2.Text = "PAINEL LABOR PLK"
        '
        'btn_Minimizar
        '
        Me.btn_Minimizar.BackgroundImage = Global.Rotinas_BK_Brasil.My.Resources.Resources.minimizar1
        Me.btn_Minimizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_Minimizar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Minimizar.FlatAppearance.BorderSize = 0
        Me.btn_Minimizar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(204, Byte), Integer), CType(CType(136, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Minimizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Minimizar.Location = New System.Drawing.Point(905, 7)
        Me.btn_Minimizar.Margin = New System.Windows.Forms.Padding(0)
        Me.btn_Minimizar.Name = "btn_Minimizar"
        Me.btn_Minimizar.Size = New System.Drawing.Size(20, 20)
        Me.btn_Minimizar.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.btn_Minimizar, "Minimizar")
        Me.btn_Minimizar.UseVisualStyleBackColor = True
        '
        'btn_Fechar
        '
        Me.btn_Fechar.BackgroundImage = Global.Rotinas_BK_Brasil.My.Resources.Resources.fechar
        Me.btn_Fechar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_Fechar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Fechar.FlatAppearance.BorderSize = 0
        Me.btn_Fechar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(171, Byte), Integer), CType(CType(28, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Fechar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Fechar.Location = New System.Drawing.Point(933, 7)
        Me.btn_Fechar.Margin = New System.Windows.Forms.Padding(0)
        Me.btn_Fechar.Name = "btn_Fechar"
        Me.btn_Fechar.Size = New System.Drawing.Size(20, 20)
        Me.btn_Fechar.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.btn_Fechar, "Fechar")
        Me.btn_Fechar.UseVisualStyleBackColor = True
        '
        'btn_Home
        '
        Me.btn_Home.BackgroundImage = Global.Rotinas_BK_Brasil.My.Resources.Resources.home
        Me.btn_Home.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_Home.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Home.FlatAppearance.BorderSize = 0
        Me.btn_Home.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(20, Byte), Integer), CType(CType(108, Byte), Integer), CType(CType(44, Byte), Integer))
        Me.btn_Home.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Home.Location = New System.Drawing.Point(7, 7)
        Me.btn_Home.Margin = New System.Windows.Forms.Padding(0)
        Me.btn_Home.Name = "btn_Home"
        Me.btn_Home.Size = New System.Drawing.Size(30, 30)
        Me.btn_Home.TabIndex = 95
        Me.ToolTip1.SetToolTip(Me.btn_Home, "Voltar para o Menu Principal")
        Me.btn_Home.UseVisualStyleBackColor = True
        '
        'form_LaborPlk
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(960, 711)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbNome_BHoras)
        Me.Controls.Add(Me.lbNome_PLbPlk)
        Me.Controls.Add(Me.btn_ArqBHoras)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbNome_Rel11Hrs)
        Me.Controls.Add(Me.btn_ArqPLbPlk)
        Me.Controls.Add(Me.lbNome_PtDia)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lbNome_EstMarc)
        Me.Controls.Add(Me.lbId_BHoras)
        Me.Controls.Add(Me.lbNome_CadFunc)
        Me.Controls.Add(Me.lb_BHoras)
        Me.Controls.Add(Me.lbNome_Rel2Hrs)
        Me.Controls.Add(Me.lbId_EstMarc)
        Me.Controls.Add(Me.btn_ArqRel2Hrs)
        Me.Controls.Add(Me.btn_ArqCadFunc)
        Me.Controls.Add(Me.lb_EstMarc)
        Me.Controls.Add(Me.btn_ArqRel11Hrs)
        Me.Controls.Add(Me.lb_Rel2Hrs)
        Me.Controls.Add(Me.btn_ArqPtDia)
        Me.Controls.Add(Me.lbId_Rel2Hrs)
        Me.Controls.Add(Me.btn_ArqEstMarc)
        Me.Controls.Add(Me.lbId_CadFunc)
        Me.Controls.Add(Me.lbId_PtDia)
        Me.Controls.Add(Me.lbId_Rel11Hrs)
        Me.Controls.Add(Me.lbId_PLbPlk)
        Me.Controls.Add(Me.lb_CadFunc)
        Me.Controls.Add(Me.lb_PtDia)
        Me.Controls.Add(Me.lb_Rel11Hrs)
        Me.Controls.Add(Me.lb_PLbPlk)
        Me.Controls.Add(Me.btn_GerarPainel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "form_LaborPlk"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Painel Labor PLK"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbNome_BHoras As Label
    Friend WithEvents btn_ArqBHoras As Button
    Friend WithEvents lbNome_Rel11Hrs As Label
    Friend WithEvents lbNome_PtDia As Label
    Friend WithEvents lbNome_EstMarc As Label
    Friend WithEvents lbNome_CadFunc As Label
    Friend WithEvents lbNome_Rel2Hrs As Label
    Friend WithEvents btn_ArqRel2Hrs As Button
    Friend WithEvents btn_ArqCadFunc As Button
    Friend WithEvents btn_ArqRel11Hrs As Button
    Friend WithEvents btn_ArqPtDia As Button
    Friend WithEvents btn_ArqEstMarc As Button
    Friend WithEvents btn_GerarPainel As Button
    Friend WithEvents lbNome_PLbPlk As Label
    Friend WithEvents btn_ArqPLbPlk As Button
    Friend WithEvents lbId_EstMarc As Label
    Friend WithEvents lb_EstMarc As Label
    Friend WithEvents lb_Rel2Hrs As Label
    Friend WithEvents lbId_Rel2Hrs As Label
    Friend WithEvents lbId_CadFunc As Label
    Friend WithEvents lbId_PtDia As Label
    Friend WithEvents lbId_Rel11Hrs As Label
    Friend WithEvents lbId_PLbPlk As Label
    Friend WithEvents lb_CadFunc As Label
    Friend WithEvents lb_PtDia As Label
    Friend WithEvents lb_Rel11Hrs As Label
    Friend WithEvents lb_PLbPlk As Label
    Friend WithEvents lbId_BHoras As Label
    Friend WithEvents lb_BHoras As Label
    Friend WithEvents ofd1 As OpenFileDialog
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents btn_Minimizar As Button
    Friend WithEvents btn_Fechar As Button
    Friend WithEvents btn_Home As Button
End Class
