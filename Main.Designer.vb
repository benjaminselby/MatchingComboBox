<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
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
        Me.MatchingComboBox1 = New temp_testing.MatchingComboBox()
        Me.SuspendLayout()
        '
        'MatchingComboBox1
        '
        Me.MatchingComboBox1.FilterRule = Nothing
        Me.MatchingComboBox1.FormattingEnabled = True
        Me.MatchingComboBox1.Location = New System.Drawing.Point(12, 12)
        Me.MatchingComboBox1.Name = "MatchingComboBox1"
        Me.MatchingComboBox1.PropertySelector = Nothing
        Me.MatchingComboBox1.Size = New System.Drawing.Size(299, 21)
        Me.MatchingComboBox1.SuggestBoxHeight = 96
        Me.MatchingComboBox1.SuggestListOrderRule = Nothing
        Me.MatchingComboBox1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.MatchingComboBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MatchingComboBox1 As MatchingComboBox
End Class
