Option Strict On
Option Infer On

Imports System.Windows.Forms
Imports System.Drawing

Partial Class Form1
    Inherits Form

    Private components As System.ComponentModel.IContainer = Nothing

    Friend WithEvents btnImport As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents lblFile As Label
    Friend WithEvents txtLog As TextBox

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then components.Dispose()
        MyBase.Dispose(disposing)
    End Sub

    Private Sub InitializeComponent()
        Me.btnImport = New Button()
        Me.btnOpen = New Button()
        Me.lblFile = New Label()
        Me.txtLog = New TextBox()
        Me.SuspendLayout()

        Me.btnImport.Location = New Point(12, 12)
        Me.btnImport.Size = New Size(150, 34)
        Me.btnImport.Text = "Import Word (.docx)"
        Me.btnImport.UseVisualStyleBackColor = True

        Me.btnOpen.Location = New Point(168, 12)
        Me.btnOpen.Size = New Size(150, 34)
        Me.btnOpen.Text = "Buka folder output"
        Me.btnOpen.Enabled = False
        Me.btnOpen.UseVisualStyleBackColor = True

        Me.lblFile.Location = New Point(12, 52)
        Me.lblFile.Size = New Size(660, 20)
        Me.lblFile.Text = "Tarik file .docx ke jendela ini, atau klik Import."

        Me.txtLog.Location = New Point(12, 78)
        Me.txtLog.Size = New Size(660, 320)
        Me.txtLog.Multiline = True
        Me.txtLog.ReadOnly = True
        Me.txtLog.ScrollBars = ScrollBars.Vertical
        Me.txtLog.Font = New Font("Consolas", 9.0!)
        Me.txtLog.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or
                           AnchorStyles.Left Or AnchorStyles.Right

        Me.AllowDrop = True
        Me.ClientSize = New Size(684, 411)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.txtLog)
        Me.Text = "RiplayWord2Html - Word ke HTML"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
