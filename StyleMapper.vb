Option Strict On
Option Infer On

Imports System.Globalization
Imports System.Text

' ============================================================================
'  StyleMapper.vb  -- konversi satuan & format Word -> deklarasi CSS
'  CssRegistry menampung deklarasi unik lalu memberi nama class pendek,
'  supaya seluruh HTML cukup memakai satu file page.css.
' ============================================================================

Public Module Units
    Public Const TwipPerCm As Double = 566.929133858
    Public Const TwipPerPt As Double = 20.0
    Public Const EmuPerPx As Double = 9525.0          ' 96 dpi

    Public Function TwToCm(tw As Integer) As Double
        Return tw / TwipPerCm
    End Function

    Public Function TwToPt(tw As Integer) As Double
        Return tw / TwipPerPt
    End Function

    Public Function EmuToPx(emu As Long) As Double
        Return emu / EmuPerPx
    End Function

    Public Function Num(v As Double) As String
        Return Math.Round(v, 3).ToString("0.###", CultureInfo.InvariantCulture)
    End Function

    Public Function Cm(tw As Integer) As String
        Return Num(TwToCm(tw)) & "cm"
    End Function

    Public Function Pt(halfPt As Integer) As String
        Return Num(halfPt / 2.0) & "pt"
    End Function
End Module

''' <summary>Kumpulan deklarasi CSS unik -> nama class.</summary>
Public Class CssRegistry
    Private ReadOnly _map As New Dictionary(Of String, String)
    Private ReadOnly _order As New List(Of String)

    ''' <summary>Daftarkan deklarasi (mis. "color:#f00;font-weight:700").
    ''' Mengembalikan nama class, atau Nothing kalau deklarasi kosong.</summary>
    Public Function ClassFor(prefix As String, decls As String) As String
        If String.IsNullOrWhiteSpace(decls) Then Return Nothing
        Dim key = prefix & "|" & decls
        Dim name As String = Nothing
        If _map.TryGetValue(key, name) Then Return name
        name = prefix & (_map.Count + 1).ToString(CultureInfo.InvariantCulture)
        _map(key) = name
        _order.Add(name & "{" & decls & "}")
        Return name
    End Function

    Public Function Render() As String
        Dim sb As New StringBuilder()
        For Each line In _order
            sb.Append("."c).AppendLine(line)
        Next
        Return sb.ToString()
    End Function
End Class

Public Module StyleMapper

    ''' <summary>Nama highlight Word -> warna CSS.</summary>
    Public Function HighlightColor(name As String) As String
        Select Case LCase(If(name, ""))
            Case "yellow" : Return "#ffff00"
            Case "green" : Return "#00ff00"
            Case "cyan" : Return "#00ffff"
            Case "turquoise" : Return "#00ffff"
            Case "magenta" : Return "#ff00ff"
            Case "blue" : Return "#0000ff"
            Case "red" : Return "#ff0000"
            Case "darkblue" : Return "#000080"
            Case "darkcyan" : Return "#008080"
            Case "darkgreen" : Return "#008000"
            Case "darkmagenta" : Return "#800080"
            Case "darkred" : Return "#800000"
            Case "darkyellow" : Return "#808000"
            Case "darkgray", "darkgrey" : Return "#808080"
            Case "lightgray", "lightgrey" : Return "#c0c0c0"
            Case "black" : Return "#000000"
            Case "white" : Return "#ffffff"
            Case Else : Return Nothing
        End Select
    End Function

    Public Function HexColor(h As String) As String
        If String.IsNullOrEmpty(h) OrElse h = "auto" Then Return Nothing
        Return "#" & h.TrimStart("#"c).ToLowerInvariant()
    End Function

    Public Function AlignCss(jc As String) As String
        Select Case LCase(If(jc, ""))
            Case "center" : Return "center"
            Case "right", "end" : Return "right"
            Case "both", "distribute" : Return "justify"
            Case "left", "start" : Return "left"
            Case Else : Return Nothing
        End Select
    End Function

    ''' <summary>Deklarasi CSS untuk satu run. `paraHalfPt` = ukuran font
    ''' yang sudah dipasang di paragraf induk, supaya tidak ditulis dua kali.</summary>
    Public Function RunDecls(r As DocRun, m As DocModel, Optional paraHalfPt As Integer = 0) As String
        Dim d As New List(Of String)
        If r.Bold Then d.Add("font-weight:700")
        If r.Italic Then d.Add("font-style:italic")
        If r.Underline Then d.Add("text-decoration:underline")

        Dim inherited = If(paraHalfPt > 0, paraHalfPt, m.DefaultSizeHalfPt)
        If r.SizeHalfPt > 0 AndAlso r.SizeHalfPt <> inherited Then
            d.Add("font-size:" & Pt(r.SizeHalfPt))
        End If
        If Not String.IsNullOrEmpty(r.FontName) AndAlso r.FontName <> m.DefaultFont Then
            d.Add("font-family:'" & r.FontName & "'")
        End If

        Dim c = HexColor(r.ColorHex)
        If c IsNot Nothing AndAlso c <> "#000000" Then d.Add("color:" & c)

        Dim hl = HighlightColor(r.HighlightName)
        If hl IsNot Nothing Then d.Add("background-color:" & hl)

        Return String.Join(";", d)
    End Function

    ''' <summary>Deklarasi CSS untuk satu paragraf.</summary>
    Public Function ParaDecls(p As DocPara, m As DocModel) As String
        Dim d As New List(Of String)

        ' Ukuran font wajib ikut di <p>: tinggi baris (strut) browser dihitung
        ' dari font-size elemen blok, bukan dari <span> di dalamnya.
        If p.SizeHalfPt > 0 AndAlso p.SizeHalfPt <> m.DefaultSizeHalfPt Then
            d.Add("font-size:" & Pt(p.SizeHalfPt))
        End If

        Dim al = AlignCss(p.Align)
        If al IsNot Nothing Then d.Add("text-align:" & al)

        ' Word: ind.left adalah indent kiri; hanging/firstLine mengubah baris pertama.
        If p.IndentLeftTw <> 0 Then d.Add("padding-left:" & Cm(p.IndentLeftTw))
        If p.IndentRightTw <> 0 Then d.Add("padding-right:" & Cm(p.IndentRightTw))
        If p.IndentHangingTw > 0 Then
            d.Add("text-indent:-" & Cm(p.IndentHangingTw))
        ElseIf p.IndentFirstLineTw <> 0 Then
            d.Add("text-indent:" & Cm(p.IndentFirstLineTw))
        End If

        If p.SpaceBeforeTw <> 0 Then d.Add("margin-top:" & Num(TwToPt(p.SpaceBeforeTw)) & "pt")
        If p.SpaceAfterTw <> 0 Then d.Add("margin-bottom:" & Num(TwToPt(p.SpaceAfterTw)) & "pt")

        If p.LineTw > 0 Then
            Select Case LCase(If(p.LineRule, "auto"))
                Case "exact", "atleast"
                    d.Add("line-height:" & Num(TwToPt(p.LineTw)) & "pt")
                Case Else
                    ' auto: w:line 240 = 1.0
                    d.Add("line-height:" & Num(p.LineTw / 240.0))
            End Select
        End If

        Dim sh = HexColor(p.ShadingHex)
        If sh IsNot Nothing Then d.Add("background-color:" & sh)

        Return String.Join(";", d)
    End Function

    Private Function BorderCss(b As DocBorder) As String
        If b Is Nothing Then Return Nothing
        If b.Val Is Nothing OrElse b.Val = "none" OrElse b.Val = "nil" Then Return "none"
        ' w:sz dalam 1/8 pt; minimal 1px supaya tetap terlihat di browser
        Dim px = Math.Max(1.0, Math.Round(b.SizeEighthPt / 8.0 * 96.0 / 72.0))
        Dim col = HexColor(b.ColorHex)
        If col Is Nothing Then col = "#000000"
        Return Num(px) & "px solid " & col
    End Function

    ''' <summary>Deklarasi CSS untuk satu sel tabel (gabung border tabel + sel).</summary>
    Public Function CellDecls(t As DocTable, c As DocCell,
                              isFirstRow As Boolean, isLastRow As Boolean,
                              isFirstCol As Boolean, isLastCol As Boolean) As String
        Dim d As New List(Of String)

        ' --- border: tcBorders menang; kalau tidak ada pakai tblBorders ---
        For Each side In {"top", "left", "bottom", "right"}
            Dim b As DocBorder = Nothing
            If c.Borders.ContainsKey(side) Then
                b = c.Borders(side)
            Else
                Dim key = side
                Select Case side
                    Case "top" : If Not isFirstRow Then key = "insideH"
                    Case "bottom" : If Not isLastRow Then key = "insideH"
                    Case "left" : If Not isFirstCol Then key = "insideV"
                    Case "right" : If Not isLastCol Then key = "insideV"
                End Select
                If t.Borders.ContainsKey(key) Then b = t.Borders(key)
            End If
            Dim css = BorderCss(b)
            If css IsNot Nothing Then d.Add("border-" & side & ":" & css)
        Next

        ' --- margin dalam sel ---
        Dim mar = If(c.MarginsTw, t.CellMarginsTw)
        If mar IsNot Nothing Then
            Dim names = {"top", "left", "bottom", "right"}
            For i = 0 To 3
                If mar(i) >= 0 Then d.Add("padding-" & names(i) & ":" & Cm(mar(i)))
            Next
        End If

        Dim sh = HexColor(c.ShadingHex)
        If sh IsNot Nothing Then d.Add("background-color:" & sh)

        Select Case LCase(If(c.VAlign, ""))
            Case "center" : d.Add("vertical-align:middle")
            Case "bottom" : d.Add("vertical-align:bottom")
            Case Else : d.Add("vertical-align:top")
        End Select

        Return String.Join(";", d)
    End Function

End Module
