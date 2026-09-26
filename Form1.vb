Option Strict On
Option Infer On

Imports System.IO
Imports System.Windows.Forms

Partial Public Class Form1

    Private _outDir As String = Nothing

    Public Sub New()
        InitializeComponent()
    End Sub

    ''' <summary>Dipanggil Program.vb untuk mode --ui &lt;docx&gt;.</summary>
    Public Sub ImportFile(path As String)
        Convert(path)
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Using dlg As New OpenFileDialog()
            dlg.Filter = "Word document (*.docx)|*.docx|Semua file (*.*)|*.*"
            dlg.Title = "Pilih file Word"
            If dlg.ShowDialog(Me) = DialogResult.OK Then Convert(dlg.FileName)
        End Using
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        If String.IsNullOrEmpty(_outDir) OrElse Not Directory.Exists(_outDir) Then Return
        Process.Start(New ProcessStartInfo(_outDir) With {.UseShellExecute = True})
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data IsNot Nothing AndAlso e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        If e.Data Is Nothing Then Return
        Dim files = TryCast(e.Data.GetData(DataFormats.FileDrop), String())
        If files Is Nothing OrElse files.Length = 0 Then Return
        Convert(files(0))
    End Sub

    Private Sub Convert(path As String)
        txtLog.Clear()
        lblFile.Text = path
        btnOpen.Enabled = False
        Application.DoEvents()

        Dim gen As New Generator()
        AddHandler gen.Log, Sub(msg) AppendLog(msg)
        Try
            _outDir = gen.Run(path)
            btnOpen.Enabled = True
            AppendLog("")
            AppendLog("OK - " & gen.BodyCount.ToString() & " body + header + footer + page.css")
        Catch ex As Exception
            AppendLog("")
            AppendLog("GAGAL: " & ex.Message)
            AppendLog(ex.StackTrace)
        End Try
    End Sub

    Private Sub AppendLog(msg As String)
        txtLog.AppendText(msg & Environment.NewLine)
        Application.DoEvents()
    End Sub

End Class
