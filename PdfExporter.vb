Option Strict On
Option Infer On

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms

' ============================================================================
'  PdfExporter.vb -- AllPages.html -> riplay.pdf
'
'  Cara utama  : WebView2 (CoreWebView2.PrintToPdfAsync) -- sama seperti
'                RiplayPdf2Html, tidak butuh browser eksternal.
'  Cadangan    : Edge / Chrome headless lewat baris perintah, dipakai otomatis
'                kalau WebView2 tidak tersedia.
'
'  page.css sudah memuat @page{size:…;margin:0} dan .page{page-break-before}
'  sehingga satu div.page = satu halaman PDF.
' ============================================================================

Public Class PdfExporter

    ''' <summary>
    ''' Cetak htmlPath ke pdfPath. Ukuran halaman dalam inci diambil dari dokumen.
    ''' Mengembalikan path PDF; melempar Exception kalau kedua cara gagal.
    ''' </summary>
    Public Shared Function Export(htmlPath As String, pdfPath As String,
                                  widthIn As Double, heightIn As Double,
                                  log As Action(Of String)) As String
        If Not File.Exists(htmlPath) Then
            Throw New FileNotFoundException("File HTML tidak ditemukan", htmlPath)
        End If
        If File.Exists(pdfPath) Then File.Delete(pdfPath)

        Try
            ExportWithWebView2(htmlPath, pdfPath, widthIn, heightIn)
        Catch ex As Exception
            log?.Invoke("  WebView2 gagal (" & ex.Message & "), pakai browser headless…")
            ExportWithBrowser(htmlPath, pdfPath)
        End Try

        If Not File.Exists(pdfPath) Then Throw New InvalidOperationException("PDF tidak terbentuk: " & pdfPath)
        log?.Invoke("PDF   : " & Path.GetFileName(pdfPath) & " (" &
                    (New FileInfo(pdfPath).Length \ 1024).ToString(CultureInfo.InvariantCulture) & " KB)")
        Return pdfPath
    End Function

    ' ------------------------------------------------------------ WebView2

    Private Shared Sub ExportWithWebView2(htmlPath As String, pdfPath As String,
                                          widthIn As Double, heightIn As Double)
        Dim err As Exception = Nothing

        Using f As New Form()
            f.ShowInTaskbar = False
            f.FormBorderStyle = FormBorderStyle.None
            f.StartPosition = FormStartPosition.Manual
            f.Location = New Drawing.Point(-3000, -3000)   ' di luar layar, tidak mengganggu
            f.Size = New Drawing.Size(1000, 1400)

            Dim web As New WebView2() With {.Dock = DockStyle.Fill}
            f.Controls.Add(web)

            AddHandler f.Shown,
                Async Sub()
                    Try
                        Dim udf = Path.Combine(Path.GetTempPath(), "RiplayWord2Html.WebView2")
                        Directory.CreateDirectory(udf)
                        Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, udf)
                        Await web.EnsureCoreWebView2Async(env)

                        Dim nav As New TaskCompletionSource(Of Boolean)
                        Dim h As EventHandler(Of CoreWebView2NavigationCompletedEventArgs) =
                            Sub(s2, e2) nav.TrySetResult(e2.IsSuccess)
                        AddHandler web.CoreWebView2.NavigationCompleted, h
                        web.CoreWebView2.Navigate("file:///" & htmlPath.Replace("\", "/"))
                        Dim ok = Await nav.Task
                        RemoveHandler web.CoreWebView2.NavigationCompleted, h
                        If Not ok Then Throw New InvalidOperationException("gagal memuat " & Path.GetFileName(htmlPath))

                        ' tunggu font selesai dimuat supaya hasil cetak sama dengan tampilan
                        For i = 1 To 50
                            Dim st = Await web.CoreWebView2.ExecuteScriptAsync("document.fonts.status")
                            If st.Contains("loaded") Then Exit For
                            Await Task.Delay(100)
                        Next

                        Dim ps = web.CoreWebView2.Environment.CreatePrintSettings()
                        ps.Orientation = CoreWebView2PrintOrientation.Portrait
                        ps.ScaleFactor = 1
                        ps.PageWidth = widthIn
                        ps.PageHeight = heightIn
                        ps.MarginTop = 0
                        ps.MarginBottom = 0
                        ps.MarginLeft = 0
                        ps.MarginRight = 0
                        ps.ShouldPrintBackgrounds = True
                        ps.ShouldPrintHeaderAndFooter = False
                        ps.ShouldPrintSelectionOnly = False

                        Dim printed = Await web.CoreWebView2.PrintToPdfAsync(pdfPath, ps)
                        If Not printed Then Throw New InvalidOperationException("PrintToPdfAsync mengembalikan False")
                    Catch ex As Exception
                        err = ex
                    Finally
                        f.Close()
                    End Try
                End Sub

            ' UI sudah punya message loop -> ShowDialog; mode CLI -> Application.Run
            If Application.MessageLoop Then
                f.ShowDialog()
            Else
                Application.Run(f)
            End If
        End Using

        If err IsNot Nothing Then Throw err
    End Sub

    ' ------------------------------------------- cadangan: browser headless

    Private Shared ReadOnly BrowserPaths As String() = {
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe"}

    Private Shared Function FindBrowser() As String
        For Each p In BrowserPaths
            If File.Exists(p) Then Return p
        Next
        Return Nothing
    End Function

    Private Shared Sub ExportWithBrowser(htmlPath As String, pdfPath As String)
        Dim exe = FindBrowser()
        If exe Is Nothing Then Throw New InvalidOperationException("Chrome/Edge tidak ditemukan untuk cetak PDF")

        Dim psi As New ProcessStartInfo(exe) With {
            .UseShellExecute = False,
            .CreateNoWindow = True,
            .RedirectStandardError = True,
            .RedirectStandardOutput = True}
        psi.ArgumentList.Add("--headless=new")
        psi.ArgumentList.Add("--disable-gpu")
        psi.ArgumentList.Add("--no-pdf-header-footer")
        psi.ArgumentList.Add("--print-to-pdf=" & pdfPath)
        psi.ArgumentList.Add("file:///" & htmlPath.Replace("\", "/"))

        Using pr = Process.Start(psi)
            pr.StandardError.ReadToEnd()
            pr.StandardOutput.ReadToEnd()
            If Not pr.WaitForExit(120000) Then
                Try
                    pr.Kill(True)
                Catch
                End Try
                Throw New TimeoutException("browser headless tidak selesai dalam 2 menit")
            End If
        End Using
    End Sub

End Class
