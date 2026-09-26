Option Strict On
Option Infer On

Imports System.IO
Imports System.Text
Imports System.Globalization

' ============================================================================
'  Generator.vb  -- orkestrasi: .docx -> page.css + header/footer/bodyN.html
'  Output: <folder docx>\Output\<nama docx>\
' ============================================================================

Public Class Generator

    Public Event Log(msg As String)

    Private Sub Say(msg As String)
        RaiseEvent Log(msg)
    End Sub

    Public Property OutputFolder As String = Nothing
    Public Property BodyCount As Integer = 0

    Public Function Run(docxPath As String) As String
        If Not File.Exists(docxPath) Then
            Throw New FileNotFoundException("File tidak ditemukan: " & docxPath)
        End If

        Dim baseDir = Path.GetDirectoryName(Path.GetFullPath(docxPath))
        Dim name = Path.GetFileNameWithoutExtension(docxPath)
        Dim outDir = Path.Combine(baseDir, "Output", name)
        OutputFolder = outDir

        Say("Baca  : " & docxPath)
        Dim reader As New DocxReader()
        Dim m = reader.Read(docxPath)
        Say(String.Format(CultureInfo.InvariantCulture,
                          "Model : {0} blok body, {1} blok header, {2} blok footer, {3} gambar",
                          m.Body.Count, m.Header.Count, m.Footer.Count, m.Media.Count))
        Say(String.Format(CultureInfo.InvariantCulture,
                          "Halaman: {0} x {1} cm, margin {2}/{3}/{4}/{5} cm",
                          Num(TwToCm(m.Setup.WidthTw)), Num(TwToCm(m.Setup.HeightTw)),
                          Num(TwToCm(m.Setup.MarginTopTw)), Num(TwToCm(m.Setup.MarginRightTw)),
                          Num(TwToCm(m.Setup.MarginBottomTw)), Num(TwToCm(m.Setup.MarginLeftTw))))

        CleanOutput(outDir)

        ' --- pecah body pada penanda <<Page Break>> ---
        Dim parts = BodySplitter.Split(m.Body)
        For Each part In parts
            BodySplitter.TrimEdgeBreaks(part)
        Next
        BodyCount = parts.Count
        Say("Pecah : " & parts.Count.ToString(CultureInfo.InvariantCulture) &
            " body (dipotong tiap menemukan <<Page Break>>)")

        ' --- tulis HTML dulu supaya CssRegistry terisi ---
        Dim css As New CssRegistry()
        Dim w As New HtmlWriter(m, css)

        Dim headerHtml = w.Document("Header", "hdr", m.Header)
        Dim footerHtml = w.Document("Footer", "ftr", m.Footer)

        Dim bodyHtml As New List(Of String)
        For i = 0 To parts.Count - 1
            bodyHtml.Add(w.Document("Body " & (i + 1).ToString(CultureInfo.InvariantCulture),
                                    "bdy", parts(i)))
        Next

        ' --- tulis semua file ---
        Directory.CreateDirectory(outDir)

        Dim cssText = BuildPageCss(m) & css.Render()
        WriteFile(Path.Combine(outDir, CssFileName), cssText)

        WriteFile(Path.Combine(outDir, "header.html"), headerHtml)
        WriteFile(Path.Combine(outDir, "footer.html"), footerHtml)
        For i = 0 To bodyHtml.Count - 1
            WriteFile(Path.Combine(outDir, "body" & (i + 1).ToString(CultureInfo.InvariantCulture) & ".html"),
                      bodyHtml(i))
        Next

        For Each kv In m.Media
            Dim mediaPath = Path.Combine(outDir, kv.Key.Replace("/"c, Path.DirectorySeparatorChar))
            Directory.CreateDirectory(Path.GetDirectoryName(mediaPath))
            File.WriteAllBytes(mediaPath, kv.Value)
            Say("Tulis : " & kv.Key)
        Next

        WriteFile(Path.Combine(outDir, "AllPages.html"),
                  BuildPreview(m, parts, w))

        Say("Selesai. Output: " & outDir)
        Return outDir
    End Function

    ''' <summary>
    ''' Gabungan header + body + footer untuk dicek di browser.
    ''' Tiap body dipecah lagi pada hard page break Word supaya jumlah halaman
    ''' preview sama dengan dokumen aslinya; {{PAGE}} diisi nomor halaman global.
    ''' </summary>
    Private Function BuildPreview(m As DocModel, parts As List(Of List(Of DocBlock)),
                                  w As HtmlWriter) As String
        Dim hdr As New StringBuilder()
        For Each b In m.Header
            w.WriteBlock(hdr, b, "      ")
        Next
        Dim ftr As New StringBuilder()
        For Each b In m.Footer
            w.WriteBlock(ftr, b, "      ")
        Next

        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang=""id"">")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset=""utf-8"">")
        sb.AppendLine("<title>AllPages</title>")
        sb.AppendLine("<link rel=""stylesheet"" href=""page.css"">")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")

        ' pecah tiap body pada hard page break -> daftar halaman Word
        Dim pages As New List(Of List(Of DocBlock))
        For Each part In parts
            pages.AddRange(BodySplitter.SplitAtPageBreaks(part))
        Next
        Dim total = pages.Count.ToString(CultureInfo.InvariantCulture)

        For i = 0 To pages.Count - 1
            Dim no = (i + 1).ToString(CultureInfo.InvariantCulture)

            Dim bdy As New StringBuilder()
            For Each b In pages(i)
                w.WriteBlock(bdy, b, "      ")
            Next

            sb.AppendLine("  <div class=""page"">")
            sb.AppendLine("    <div class=""hdr"">")
            sb.Append(Fill(hdr.ToString(), no, total))
            sb.AppendLine("    </div>")
            sb.AppendLine("    <div class=""bdy"">")
            sb.Append(bdy.ToString())
            sb.AppendLine("    </div>")
            sb.AppendLine("    <div class=""ftr"">")
            sb.Append(Fill(ftr.ToString(), no, total))
            sb.AppendLine("    </div>")
            sb.AppendLine("  </div>")
        Next
        Say("Preview: AllPages.html = " & total & " halaman")

        sb.AppendLine("</body>")
        sb.AppendLine("</html>")
        Return sb.ToString()
    End Function

    Private Shared Function Fill(html As String, page As String, total As String) As String
        Return html.Replace("{{PAGE}}", page).Replace("{{PAGES}}", total)
    End Function

    Private Sub WriteFile(filePath As String, text As String)
        File.WriteAllText(filePath, text, New UTF8Encoding(False))
        Say("Tulis : " & Path.GetFileName(filePath))
    End Sub

    Private Sub CleanOutput(dir As String)
        If Not Directory.Exists(dir) Then Return
        For Each f In Directory.GetFiles(dir, "*.html")
            File.Delete(f)
        Next
        Dim cssPath = Path.Combine(dir, CssFileName)
        If File.Exists(cssPath) Then File.Delete(cssPath)
        Say("Bersih: output lama dihapus")
    End Sub

End Class
