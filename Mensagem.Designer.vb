<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_Mensagem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(form_Mensagem))
        Me.lb_Titulo = New System.Windows.Forms.Label()
        Me.lb_Escolha = New System.Windows.Forms.Label()
        Me.btn_Nao = New System.Windows.Forms.Button()
        Me.btn_Sim = New System.Windows.Forms.Button()
        Me.lb_Aviso = New System.Windows.Forms.Label()
        Me.pb_Estilo = New System.Windows.Forms.PictureBox()
        Me.p_Titulo = New System.Windows.Forms.Panel()
        Me.btn_Ok = New System.Windows.Forms.Button()
        CType(Me.pb_Estilo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.p_Titulo.SuspendLayout()
        Me.SuspendLayout()
        '
        'lb_Titulo
        '
        Me.lb_Titulo.AutoSize = True
        Me.lb_Titulo.Font = New System.Drawing.Font("Cooper Black", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lb_Titulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.lb_Titulo.Location = New System.Drawing.Point(40, 35)
        Me.lb_Titulo.MinimumSize = New System.Drawing.Size(700, 0)
        Me.lb_Titulo.Name = "lb_Titulo"
        Me.lb_Titulo.Size = New System.Drawing.Size(700, 32)
        Me.lb_Titulo.TabIndex = 0
        Me.lb_Titulo.Text = "Label1"
        Me.lb_Titulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lb_Escolha
        '
        Me.lb_Escolha.AutoSize = True
        Me.lb_Escolha.Location = New System.Drawing.Point(717, 128)
        Me.lb_Escolha.Name = "lb_Escolha"
        Me.lb_Escolha.Size = New System.Drawing.Size(51, 17)
        Me.lb_Escolha.TabIndex = 12
        Me.lb_Escolha.Text = "Label1"
        Me.lb_Escolha.Visible = False
        '
        'btn_Nao
        '
        Me.btn_Nao.BackColor = System.Drawing.Color.FromArgb(CType(CType(80, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(20, Byte), Integer))
        Me.btn_Nao.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Nao.FlatAppearance.BorderSize = 0
        Me.btn_Nao.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(80, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(20, Byte), Integer))
        Me.btn_Nao.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(28, Byte), Integer), CType(CType(16, Byte), Integer))
        Me.btn_Nao.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Nao.Font = New System.Drawing.Font("Cooper Black", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Nao.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_Nao.Location = New System.Drawing.Point(25, 350)
        Me.btn_Nao.Name = "btn_Nao"
        Me.btn_Nao.Size = New System.Drawing.Size(350, 65)
        Me.btn_Nao.TabIndex = 11
        Me.btn_Nao.Text = "Não"
        Me.btn_Nao.UseVisualStyleBackColor = False
        '
        'btn_Sim
        '
        Me.btn_Sim.BackColor = System.Drawing.Color.FromArgb(CType(CType(215, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Sim.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Sim.FlatAppearance.BorderSize = 0
        Me.btn_Sim.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(215, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Sim.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(171, Byte), Integer), CType(CType(28, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Sim.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Sim.Font = New System.Drawing.Font("Cooper Black", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Sim.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_Sim.Location = New System.Drawing.Point(400, 350)
        Me.btn_Sim.Name = "btn_Sim"
        Me.btn_Sim.Size = New System.Drawing.Size(350, 65)
        Me.btn_Sim.TabIndex = 10
        Me.btn_Sim.Text = "Sim"
        Me.btn_Sim.UseVisualStyleBackColor = False
        '
        'lb_Aviso
        '
        Me.lb_Aviso.AutoSize = True
        Me.lb_Aviso.Font = New System.Drawing.Font("Chicken Sans", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lb_Aviso.Location = New System.Drawing.Point(212, 125)
        Me.lb_Aviso.MinimumSize = New System.Drawing.Size(500, 200)
        Me.lb_Aviso.Name = "lb_Aviso"
        Me.lb_Aviso.Size = New System.Drawing.Size(500, 200)
        Me.lb_Aviso.TabIndex = 9
        Me.lb_Aviso.Text = "Label1"
        Me.lb_Aviso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pb_Estilo
        '
        Me.pb_Estilo.Location = New System.Drawing.Point(25, 150)
        Me.pb_Estilo.Name = "pb_Estilo"
        Me.pb_Estilo.Size = New System.Drawing.Size(150, 150)
        Me.pb_Estilo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pb_Estilo.TabIndex = 8
        Me.pb_Estilo.TabStop = False
        '
        'p_Titulo
        '
        Me.p_Titulo.Controls.Add(Me.lb_Titulo)
        Me.p_Titulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.p_Titulo.Location = New System.Drawing.Point(0, 0)
        Me.p_Titulo.Name = "p_Titulo"
        Me.p_Titulo.Size = New System.Drawing.Size(780, 100)
        Me.p_Titulo.TabIndex = 7
        '
        'btn_Ok
        '
        Me.btn_Ok.BackColor = System.Drawing.Color.FromArgb(CType(CType(80, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(20, Byte), Integer))
        Me.btn_Ok.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btn_Ok.FlatAppearance.BorderSize = 0
        Me.btn_Ok.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(80, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(20, Byte), Integer))
        Me.btn_Ok.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(28, Byte), Integer), CType(CType(16, Byte), Integer))
        Me.btn_Ok.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Ok.Font = New System.Drawing.Font("Cooper Black", 16.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Ok.ForeColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btn_Ok.Location = New System.Drawing.Point(25, 350)
        Me.btn_Ok.Name = "btn_Ok"
        Me.btn_Ok.Size = New System.Drawing.Size(725, 65)
        Me.btn_Ok.TabIndex = 13
        Me.btn_Ok.Text = "OK"
        Me.btn_Ok.UseVisualStyleBackColor = False
        '
        'form_Mensagem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(780, 447)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Ok)
        Me.Controls.Add(Me.lb_Escolha)
        Me.Controls.Add(Me.btn_Nao)
        Me.Controls.Add(Me.btn_Sim)
        Me.Controls.Add(Me.lb_Aviso)
        Me.Controls.Add(Me.pb_Estilo)
        Me.Controls.Add(Me.p_Titulo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "form_Mensagem"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aviso"
        CType(Me.pb_Estilo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.p_Titulo.ResumeLayout(False)
        Me.p_Titulo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lb_Titulo As Label
    Friend WithEvents lb_Escolha As Label
    Friend WithEvents btn_Nao As Button
    Friend WithEvents btn_Sim As Button
    Friend WithEvents lb_Aviso As Label
    Friend WithEvents pb_Estilo As PictureBox
    Friend WithEvents p_Titulo As Panel
    Friend WithEvents btn_Ok As Button
End Class
