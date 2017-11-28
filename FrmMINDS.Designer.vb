<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMINDS
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.dtpProcesar1 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnPagos = New System.Windows.Forms.Button()
        Me.btnLCtos = New System.Windows.Forms.Button()
        Me.btnCliente = New System.Windows.Forms.Button()
        Me.BttPromo = New System.Windows.Forms.Button()
        Me.dtpProcesar2 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'dtpProcesar1
        '
        Me.dtpProcesar1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpProcesar1.Location = New System.Drawing.Point(58, 10)
        Me.dtpProcesar1.Name = "dtpProcesar1"
        Me.dtpProcesar1.Size = New System.Drawing.Size(88, 20)
        Me.dtpProcesar1.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Desde"
        '
        'btnPagos
        '
        Me.btnPagos.ForeColor = System.Drawing.Color.Black
        Me.btnPagos.Location = New System.Drawing.Point(161, 32)
        Me.btnPagos.Name = "btnPagos"
        Me.btnPagos.Size = New System.Drawing.Size(117, 23)
        Me.btnPagos.TabIndex = 18
        Me.btnPagos.Text = "Carga Pagos"
        Me.btnPagos.UseVisualStyleBackColor = True
        '
        'btnLCtos
        '
        Me.btnLCtos.ForeColor = System.Drawing.Color.Black
        Me.btnLCtos.Location = New System.Drawing.Point(24, 112)
        Me.btnLCtos.Name = "btnLCtos"
        Me.btnLCtos.Size = New System.Drawing.Size(117, 23)
        Me.btnLCtos.TabIndex = 19
        Me.btnLCtos.Text = "Layout Contratos"
        Me.btnLCtos.UseVisualStyleBackColor = True
        Me.btnLCtos.Visible = False
        '
        'btnCliente
        '
        Me.btnCliente.ForeColor = System.Drawing.Color.Black
        Me.btnCliente.Location = New System.Drawing.Point(24, 83)
        Me.btnCliente.Name = "btnCliente"
        Me.btnCliente.Size = New System.Drawing.Size(117, 23)
        Me.btnCliente.TabIndex = 21
        Me.btnCliente.Text = "Layout Clientes"
        Me.btnCliente.UseVisualStyleBackColor = True
        Me.btnCliente.Visible = False
        '
        'BttPromo
        '
        Me.BttPromo.ForeColor = System.Drawing.Color.Black
        Me.BttPromo.Location = New System.Drawing.Point(24, 141)
        Me.BttPromo.Name = "BttPromo"
        Me.BttPromo.Size = New System.Drawing.Size(117, 23)
        Me.BttPromo.TabIndex = 22
        Me.BttPromo.Text = "Carga Promotores"
        Me.BttPromo.UseVisualStyleBackColor = True
        Me.BttPromo.Visible = False
        '
        'dtpProcesar2
        '
        Me.dtpProcesar2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpProcesar2.Location = New System.Drawing.Point(58, 36)
        Me.dtpProcesar2.Name = "dtpProcesar2"
        Me.dtpProcesar2.Size = New System.Drawing.Size(88, 20)
        Me.dtpProcesar2.TabIndex = 23
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Hasta"
        '
        'FrmMINDS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(300, 72)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpProcesar2)
        Me.Controls.Add(Me.BttPromo)
        Me.Controls.Add(Me.btnCliente)
        Me.Controls.Add(Me.btnLCtos)
        Me.Controls.Add(Me.btnPagos)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpProcesar1)
        Me.Name = "FrmMINDS"
        Me.Text = "Carga de Pagos "
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpProcesar1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnPagos As System.Windows.Forms.Button
    Friend WithEvents btnLCtos As System.Windows.Forms.Button
    Friend WithEvents btnCliente As System.Windows.Forms.Button
    Friend WithEvents BttPromo As System.Windows.Forms.Button
    Friend WithEvents dtpProcesar2 As DateTimePicker
    Friend WithEvents Label2 As Label
End Class
