Option Strict On
Option Infer On

Imports System.Globalization
Imports System.Text

' ============================================================================
'  HtmlWriter.vb  -- DocBlock -> HTML, memakai CssRegistry supaya semua gaya
'  mendarat di satu file page.css.
' ============================================================================

Public Class HtmlWriter

    Private ReadOnly _m As DocModel
    Private ReadOnly _css As CssRegistry

    Public Sub New(model As DocModel, css As CssRegistry)
        _m = model
        _css = css
    End Sub

    ' ------------------------------------------------------------- utilitas

    Public Shared Function Esc(s As String) As String
        If s Is Nothing Then Return ""
        Dim sb As New StringBuilder(s.Length)
        For Each ch In s
            Select Case ch
                Case "&"c : sb.Append("&amp;")
                Case "<"c : sb.Append("&lt;")
                Case ">"c : sb.Append("&gt;")
                Case ChrW(160) : sb.Append("&nbsp;")
                Case Else : sb.Append(ch)
            End Select
        Next
        ' jaga spasi ganda supaya tidak diciutkan browser
        Return sb.ToString().Replace("  ", " &nbsp;")
    End Function

    Private Function Cls(prefix As String, decls As String) As String
        Dim c = _css.ClassFor(prefix, decls)
        If c Is Nothing Then Return ""
        Return " class=""" & c & """"
    End Function

    ' ------------------------------------------------------------------ run

    Private Sub WriteRun(sb As StringBuilder, r As DocRun, paraHalfPt As Integer)
        Select Case r.Kind
            Case RunKind.LineBreak
                sb.Append("<br>")

            Case RunKind.PageBreak
                ' ditangani di level paragraf

            Case RunKind.Tab
                sb.Append("<span class=""tab""></span>")

            Case RunKind.Picture
                Dim w = EmuToPx(r.WidthEmu)
                Dim h = EmuToPx(r.HeightEmu)
                sb.Append("<img src=""").Append(r.ImageFile).Append(""" alt="""" style=""width:").
                   Append(Num(w)).Append("px;height:").Append(Num(h)).Append("px"">")

            Case Else
                Dim txt = Esc(r.Text)
                If txt.Length = 0 Then Return
                Dim decls = RunDecls(r, _m, paraHalfPt)
                Dim tag = "span"
                If r.Superscript Then
                    tag = "sup"
                ElseIf r.Subscript Then
                    tag = "sub"
                End If
                If decls.Length = 0 AndAlso tag = "span" Then
                    sb.Append(txt)
                Else
                    sb.Append("<"c).Append(tag).Append(Cls("r", decls)).Append(">"c).
                       Append(txt).Append("</").Append(tag).Append(">"c)
                End If
        End Select
    End Sub

    ' ------------------------------------------------------------- paragraf

    ''' <summary>
    ''' Paragraf hanging-indent Word memakai tab stop: "1.<tab>teks".
    ''' Di HTML label dibungkus inline-block selebar hanging, persis seperti
    ''' perilaku tab stop Word, lalu tab-nya dibuang.
    ''' </summary>
    Private Function WriteHangingLabel(sb As StringBuilder, p As DocPara) As Integer
        Dim paraHalfPt = p.SizeHalfPt
        If p.IndentHangingTw <= 0 Then Return 0
        Dim tabAt = -1
        For i = 0 To p.Runs.Count - 1
            If p.Runs(i).Kind = RunKind.Tab Then
                tabAt = i
                Exit For
            End If
        Next
        If tabAt < 0 Then Return 0

        sb.Append("<span class=""lbl"" style=""width:").Append(Cm(p.IndentHangingTw)).Append(""">")
        For i = 0 To tabAt - 1
            WriteRun(sb, p.Runs(i), paraHalfPt)
        Next
        sb.Append("</span>")
        Return tabAt + 1
    End Function

    Public Sub WritePara(sb As StringBuilder, p As DocPara, indent As String)
        If p.IsPageBreakOnly Then
            sb.Append(indent).AppendLine("<div class=""page-break""></div>")
            Return
        End If

        Dim paraHalfPt = p.SizeHalfPt
        Dim decls = ParaDecls(p, _m)
        sb.Append(indent).Append("<p").Append(Cls("p", decls)).Append(">"c)

        Dim start = WriteHangingLabel(sb, p)
        Dim empty = True
        For i = start To p.Runs.Count - 1
            Dim r = p.Runs(i)
            If r.Kind = RunKind.PageBreak Then Continue For
            WriteRun(sb, r, paraHalfPt)
            If r.Kind <> RunKind.Tab Then empty = False
        Next

        ' Paragraf kosong tetap punya tinggi di Word sesuai ukuran font run-nya.
        ' Tanpa ini &nbsp; akan mewarisi font-size default dan barisnya jadi terlalu tinggi.
        If empty AndAlso start = 0 Then sb.Append("&nbsp;")

        sb.AppendLine("</p>")
    End Sub

    ' ---------------------------------------------------------------- tabel

    Public Sub WriteTable(sb As StringBuilder, t As DocTable, indent As String)
        Dim totalTw = 0
        For Each g In t.GridTw
            totalTw += g
        Next

        Dim tDecls As New List(Of String)
        If totalTw > 0 Then tDecls.Add("width:" & Cm(totalTw))
        If t.IndentTw <> 0 Then tDecls.Add("margin-left:" & Cm(t.IndentTw))

        sb.Append(indent).Append("<table").Append(Cls("t", String.Join(";", tDecls))).AppendLine(">")

        If t.GridTw.Count > 0 Then
            sb.Append(indent).AppendLine("  <colgroup>")
            For Each g In t.GridTw
                sb.Append(indent).Append("    <col style=""width:").Append(Cm(g)).AppendLine(""">")
            Next
            sb.Append(indent).AppendLine("  </colgroup>")
        End If

        For ri = 0 To t.Rows.Count - 1
            Dim row = t.Rows(ri)
            Dim rDecls As New List(Of String)
            If row.HeightTw > 0 Then
                ' <tr> mengabaikan min-height, jadi pakai height:
                ' browser tetap melebarkan baris kalau isinya lebih tinggi (= trHeight atLeast)
                rDecls.Add("height:" & Cm(row.HeightTw))
            End If
            sb.Append(indent).Append("  <tr").Append(Cls("tr", String.Join(";", rDecls))).AppendLine(">")

            Dim colIdx = 0
            For ci = 0 To row.Cells.Count - 1
                Dim c = row.Cells(ci)
                Dim span = Math.Max(1, c.GridSpan)
                If c.Skip Then
                    colIdx += span
                    Continue For
                End If

                Dim isFirstCol = (colIdx = 0)
                Dim isLastCol = (colIdx + span >= t.GridTw.Count)
                Dim decls = CellDecls(t, c, ri = 0, ri = t.Rows.Count - 1, isFirstCol, isLastCol)

                sb.Append(indent).Append("    <td")
                If span > 1 Then sb.Append(" colspan=""").Append(span.ToString(CultureInfo.InvariantCulture)).Append(""""c)
                If c.RowSpan > 1 Then sb.Append(" rowspan=""").Append(c.RowSpan.ToString(CultureInfo.InvariantCulture)).Append(""""c)
                sb.Append(Cls("c", decls)).AppendLine(">")

                For Each b In c.Blocks
                    WriteBlock(sb, b, indent & "      ")
                Next

                sb.Append(indent).AppendLine("    </td>")
                colIdx += span
            Next

            sb.Append(indent).AppendLine("  </tr>")
        Next

        sb.Append(indent).AppendLine("</table>")
    End Sub

    Public Sub WriteBlock(sb As StringBuilder, b As DocBlock, indent As String)
        If b.IsPara Then
            WritePara(sb, b.Para, indent)
        ElseIf b.IsTable Then
            WriteTable(sb, b.Table, indent)
        End If
    End Sub

    ''' <summary>Bungkus daftar blok jadi satu dokumen HTML lengkap.</summary>
    Public Function Document(title As String, wrapperClass As String,
                             blocks As List(Of DocBlock)) As String
        Dim inner As New StringBuilder()
        For Each b In blocks
            WriteBlock(inner, b, "    ")
        Next

        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang=""id"">")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset=""utf-8"">")
        sb.Append("<title>").Append(Esc(title)).AppendLine("</title>")
        sb.AppendLine("<link rel=""stylesheet"" href=""page.css"">")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.Append("  <div class=""").Append(wrapperClass).AppendLine(""">")
        sb.Append(inner.ToString())
        sb.AppendLine("  </div>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")
        Return sb.ToString()
    End Function

End Class
