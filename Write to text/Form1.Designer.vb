<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Encode = New System.Windows.Forms.Button()
        Me.Entry = New System.Windows.Forms.TextBox()
        Me.Decode = New System.Windows.Forms.Button()
        Me.Key = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Encode
        '
        Me.Encode.Location = New System.Drawing.Point(13, 64)
        Me.Encode.Name = "Encode"
        Me.Encode.Size = New System.Drawing.Size(124, 76)
        Me.Encode.TabIndex = 0
        Me.Encode.Text = "Encode"
        Me.Encode.UseVisualStyleBackColor = True
        '
        'Entry
        '
        Me.Entry.Location = New System.Drawing.Point(12, 12)
        Me.Entry.Name = "Entry"
        Me.Entry.Size = New System.Drawing.Size(260, 20)
        Me.Entry.TabIndex = 1
        Me.Entry.Text = "Enter the message..."
        '
        'Decode
        '
        Me.Decode.Location = New System.Drawing.Point(148, 64)
        Me.Decode.Name = "Decode"
        Me.Decode.Size = New System.Drawing.Size(124, 76)
        Me.Decode.TabIndex = 2
        Me.Decode.Text = "Decode"
        Me.Decode.UseVisualStyleBackColor = True
        '
        'Key
        '
        Me.Key.Location = New System.Drawing.Point(12, 38)
        Me.Key.Name = "Key"
        Me.Key.Size = New System.Drawing.Size(257, 20)
        Me.Key.TabIndex = 3
        Me.Key.Text = "Key"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 149)
        Me.Controls.Add(Me.Key)
        Me.Controls.Add(Me.Decode)
        Me.Controls.Add(Me.Entry)
        Me.Controls.Add(Me.Encode)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "EN/DECRYPT"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Encode As Button
    Friend WithEvents Entry As TextBox
    Friend WithEvents Decode As Button
    Friend WithEvents Key As TextBox
End Class
