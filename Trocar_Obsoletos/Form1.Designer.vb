<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button_ok = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_pasta_prod = New System.Windows.Forms.TextBox()
        Me.Button_browse = New System.Windows.Forms.Button()
        Me.Button_Limpar = New System.Windows.Forms.Button()
        Me.Button_sair = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Desenho = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Localizacao = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Extensao = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Obsoleto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Mover = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NOME_ANTIGO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowDrop = True
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Desenho, Me.Localizacao, Me.Extensao, Me.Obsoleto, Me.Mover, Me.NOME_ANTIGO})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(640, 430)
        Me.DataGridView1.TabIndex = 0
        '
        'Button_ok
        '
        Me.Button_ok.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Button_ok.BackColor = System.Drawing.SystemColors.Control
        Me.Button_ok.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlDark
        Me.Button_ok.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Button_ok.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_ok.Location = New System.Drawing.Point(295, 507)
        Me.Button_ok.Name = "Button_ok"
        Me.Button_ok.Size = New System.Drawing.Size(75, 23)
        Me.Button_ok.TabIndex = 1
        Me.Button_ok.Text = "OK"
        Me.Button_ok.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 458)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Pasta na produção:"
        '
        'TextBox_pasta_prod
        '
        Me.TextBox_pasta_prod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox_pasta_prod.Location = New System.Drawing.Point(116, 454)
        Me.TextBox_pasta_prod.Name = "TextBox_pasta_prod"
        Me.TextBox_pasta_prod.Size = New System.Drawing.Size(471, 20)
        Me.TextBox_pasta_prod.TabIndex = 3
        '
        'Button_browse
        '
        Me.Button_browse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_browse.BackColor = System.Drawing.SystemColors.Control
        Me.Button_browse.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlDark
        Me.Button_browse.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Button_browse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_browse.Location = New System.Drawing.Point(591, 453)
        Me.Button_browse.Name = "Button_browse"
        Me.Button_browse.Size = New System.Drawing.Size(61, 23)
        Me.Button_browse.TabIndex = 4
        Me.Button_browse.Text = "Browse"
        Me.Button_browse.UseVisualStyleBackColor = False
        '
        'Button_Limpar
        '
        Me.Button_Limpar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Limpar.BackColor = System.Drawing.SystemColors.Control
        Me.Button_Limpar.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlDark
        Me.Button_Limpar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button_Limpar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Limpar.Location = New System.Drawing.Point(518, 507)
        Me.Button_Limpar.Name = "Button_Limpar"
        Me.Button_Limpar.Size = New System.Drawing.Size(64, 23)
        Me.Button_Limpar.TabIndex = 5
        Me.Button_Limpar.Text = "Limpar"
        Me.Button_Limpar.UseVisualStyleBackColor = False
        '
        'Button_sair
        '
        Me.Button_sair.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_sair.BackColor = System.Drawing.SystemColors.Control
        Me.Button_sair.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlDark
        Me.Button_sair.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Red
        Me.Button_sair.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_sair.Location = New System.Drawing.Point(588, 507)
        Me.Button_sair.Name = "Button_sair"
        Me.Button_sair.Size = New System.Drawing.Size(64, 23)
        Me.Button_sair.TabIndex = 6
        Me.Button_sair.Text = "Sair"
        Me.Button_sair.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.Button_sair.UseVisualStyleBackColor = False
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.AddExtension = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 480)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(640, 20)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar1.TabIndex = 7
        Me.ProgressBar1.Visible = False
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(315, 484)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 12)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Desenho
        '
        Me.Desenho.HeaderText = "Desenho"
        Me.Desenho.Name = "Desenho"
        Me.Desenho.ReadOnly = True
        Me.Desenho.Width = 75
        '
        'Localizacao
        '
        Me.Localizacao.HeaderText = "Localização"
        Me.Localizacao.Name = "Localizacao"
        Me.Localizacao.ReadOnly = True
        Me.Localizacao.Visible = False
        Me.Localizacao.Width = 89
        '
        'Extensao
        '
        Me.Extensao.HeaderText = "Extensão"
        Me.Extensao.Name = "Extensao"
        Me.Extensao.ReadOnly = True
        Me.Extensao.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Extensao.Width = 76
        '
        'Obsoleto
        '
        Me.Obsoleto.HeaderText = "Obsoleto"
        Me.Obsoleto.Name = "Obsoleto"
        Me.Obsoleto.ReadOnly = True
        Me.Obsoleto.ToolTipText = "Tornar ficheiro antigo obsoleto"
        Me.Obsoleto.Width = 74
        '
        'Mover
        '
        Me.Mover.HeaderText = "Mover"
        Me.Mover.Name = "Mover"
        Me.Mover.ReadOnly = True
        Me.Mover.ToolTipText = "Mover novo ficheiro para a pasta da produção"
        Me.Mover.Width = 62
        '
        'NOME_ANTIGO
        '
        Me.NOME_ANTIGO.HeaderText = "NOME_ANTIGO"
        Me.NOME_ANTIGO.Name = "NOME_ANTIGO"
        Me.NOME_ANTIGO.Visible = False
        Me.NOME_ANTIGO.Width = 111
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(664, 542)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button_sair)
        Me.Controls.Add(Me.Button_Limpar)
        Me.Controls.Add(Me.Button_browse)
        Me.Controls.Add(Me.TextBox_pasta_prod)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button_ok)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(589, 300)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "I&D - Trocar desenhos obsoletos"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button_ok As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox_pasta_prod As TextBox
    Friend WithEvents Button_browse As Button
    Friend WithEvents Button_Limpar As Button
    Friend WithEvents Button_sair As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label2 As Label
    Friend WithEvents Desenho As DataGridViewTextBoxColumn
    Friend WithEvents Localizacao As DataGridViewTextBoxColumn
    Friend WithEvents Extensao As DataGridViewTextBoxColumn
    Friend WithEvents Obsoleto As DataGridViewTextBoxColumn
    Friend WithEvents Mover As DataGridViewTextBoxColumn
    Friend WithEvents NOME_ANTIGO As DataGridViewTextBoxColumn
End Class
