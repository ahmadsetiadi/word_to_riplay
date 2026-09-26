Option Strict On
Option Infer On

Imports System.Windows.Forms

' ============================================================================
'  Program.vb  -- entry point.
'    RiplayWord2Html.exe                      -> UI
'    RiplayWord2Html.exe --ui <file.docx>     -> UI + langsung konversi
'    RiplayWord2Html.exe <file.docx>          -> CLI, tanpa UI
' ============================================================================

Public Module Program

    <STAThread>
    Public Sub Main(args As String())
        Dim docx As String = Nothing
        Dim useUi = True

        If args.Length > 0 Then
            If String.Equals(args(0), "--ui", StringComparison.OrdinalIgnoreCase) Then
                If args.Length > 1 Then docx = args(1)
            Else
                docx = args(0)
                useUi = False
            End If
        End If

        If Not useUi Then
            Environment.ExitCode = RunCli(docx)
            Return
        End If

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Dim f As New Form1()
        If Not String.IsNullOrEmpty(docx) Then
            AddHandler f.Shown, Sub(s, e) f.ImportFile(docx)
        End If
        Application.Run(f)
    End Sub

    Private Function RunCli(docx As String) As Integer
        Dim gen As New Generator()
        AddHandler gen.Log, Sub(msg) Console.WriteLine(msg)
        Try
            gen.Run(docx)
            Return 0
        Catch ex As Exception
            Console.Error.WriteLine("GAGAL: " & ex.Message)
            Console.Error.WriteLine(ex.StackTrace)
            Return 1
        End Try
    End Function

End Module
