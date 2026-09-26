Option Strict On
Option Infer On

' ============================================================================
'  Models.vb  -- struktur data hasil baca .docx
'  Satuan disimpan apa adanya dari Word:
'    - twip / dxa  : 1/1440 inci  (indent, margin, lebar tabel)
'    - half-point  : 1/2 pt       (ukuran font w:sz)
'    - eighth-point: 1/8 pt       (tebal border w:sz)
'    - EMU         : 1/914400 inci (ukuran gambar)
'  Konversi ke CSS dikerjakan StyleMapper.vb, bukan di sini.
' ============================================================================

Public Enum RunKind
    Text        ' teks biasa
    Tab         ' <w:tab/>
    LineBreak   ' <w:br/>
    PageBreak   ' <w:br w:type="page"/>
    Picture     ' <w:drawing>
    Field       ' hasil field PAGE / NUMPAGES -> placeholder
End Enum

''' <summary>Satu potong teks dengan format seragam.</summary>
Public Class DocRun
    Public Property Kind As RunKind = RunKind.Text
    Public Property Text As String = ""

    Public Property Bold As Boolean
    Public Property Italic As Boolean
    Public Property Underline As Boolean
    Public Property Superscript As Boolean
    Public Property Subscript As Boolean

    Public Property FontName As String = Nothing
    Public Property SizeHalfPt As Integer = 0          ' 0 = ikut default
    Public Property ColorHex As String = Nothing       ' "FF0000"
    Public Property HighlightName As String = Nothing  ' "turquoise", "yellow", ...

    ' khusus Kind = Picture
    Public Property ImageFile As String = Nothing      ' nama file relatif, mis. "media/image1.png"
    Public Property WidthEmu As Long = 0
    Public Property HeightEmu As Long = 0
End Class

''' <summary>Satu paragraf.</summary>
Public Class DocPara
    Public Property Runs As New List(Of DocRun)

    Public Property Align As String = Nothing          ' left/center/right/both
    Public Property IndentLeftTw As Integer = 0
    Public Property IndentRightTw As Integer = 0
    Public Property IndentFirstLineTw As Integer = 0   ' positif = first-line indent
    Public Property IndentHangingTw As Integer = 0     ' positif = hanging
    Public Property SpaceBeforeTw As Integer = 0
    Public Property SpaceAfterTw As Integer = 0
    Public Property LineTw As Integer = 0              ' w:line
    Public Property LineRule As String = Nothing       ' auto / exact / atLeast
    Public Property ShadingHex As String = Nothing     ' w:shd w:fill
    Public Property TabStopsTw As New List(Of Integer)
    Public Property KeepNext As Boolean
    ''' <summary>w:pPr/w:rPr/w:sz -- ukuran tanda paragraf.</summary>
    Public Property MarkSizeHalfPt As Integer = 0

    ''' <summary>
    ''' Ukuran font yang menentukan tinggi "strut" baris di browser.
    ''' Diambil dari tanda paragraf, kalau tidak ada dari run teks pertama.
    ''' </summary>
    Public ReadOnly Property SizeHalfPt As Integer
        Get
            If MarkSizeHalfPt > 0 Then Return MarkSizeHalfPt
            For Each r In Runs
                If r.SizeHalfPt > 0 Then Return r.SizeHalfPt
            Next
            Return 0
        End Get
    End Property

    Public ReadOnly Property PlainText As String
        Get
            Dim sb As New System.Text.StringBuilder()
            For Each r In Runs
                If r.Kind = RunKind.Text OrElse r.Kind = RunKind.Field Then sb.Append(r.Text)
            Next
            Return sb.ToString()
        End Get
    End Property

    ''' <summary>True kalau paragraf ini isinya cuma hard page break.</summary>
    Public ReadOnly Property IsPageBreakOnly As Boolean
        Get
            Dim hasBreak = False
            For Each r In Runs
                Select Case r.Kind
                    Case RunKind.PageBreak
                        hasBreak = True
                    Case RunKind.Text
                        If r.Text.Trim().Length > 0 Then Return False
                    Case Else
                        Return False
                End Select
            Next
            Return hasBreak
        End Get
    End Property
End Class

Public Class DocCell
    Public Property Blocks As New List(Of DocBlock)
    Public Property WidthTw As Integer = 0
    Public Property GridSpan As Integer = 1
    Public Property VMerge As String = Nothing         ' "restart" / "continue"
    Public Property VAlign As String = Nothing         ' top / center / bottom
    Public Property ShadingHex As String = Nothing
    Public Property Borders As New Dictionary(Of String, DocBorder)  ' top/left/bottom/right
    Public Property MarginsTw As Integer() = Nothing   ' {top, left, bottom, right}

    ' diisi saat penulisan HTML
    Public Property RowSpan As Integer = 1
    Public Property Skip As Boolean = False
End Class

Public Class DocBorder
    Public Property Val As String = "none"             ' single / none / ...
    Public Property SizeEighthPt As Integer = 0
    Public Property ColorHex As String = "auto"
End Class

Public Class DocRow
    Public Property Cells As New List(Of DocCell)
    Public Property HeightTw As Integer = 0
    Public Property HeightRule As String = Nothing     ' atLeast / exact
    Public Property CantSplit As Boolean
End Class

Public Class DocTable
    Public Property Rows As New List(Of DocRow)
    Public Property GridTw As New List(Of Integer)
    Public Property IndentTw As Integer = 0
    Public Property Borders As New Dictionary(Of String, DocBorder)  ' + insideH / insideV
    Public Property CellMarginsTw As Integer() = Nothing              ' {top, left, bottom, right}
End Class

''' <summary>Blok level atas: paragraf ATAU tabel.</summary>
Public Class DocBlock
    Public Property Para As DocPara = Nothing
    Public Property Table As DocTable = Nothing
    Public ReadOnly Property IsPara As Boolean
        Get
            Return Para IsNot Nothing
        End Get
    End Property
    Public ReadOnly Property IsTable As Boolean
        Get
            Return Table IsNot Nothing
        End Get
    End Property

    Public Shared Function FromPara(p As DocPara) As DocBlock
        Return New DocBlock With {.Para = p}
    End Function

    Public Shared Function FromTable(t As DocTable) As DocBlock
        Return New DocBlock With {.Table = t}
    End Function
End Class

''' <summary>Geometri halaman dari w:sectPr (satuan twip).</summary>
Public Class PageSetup
    Public Property WidthTw As Integer = 11906
    Public Property HeightTw As Integer = 16838
    Public Property MarginTopTw As Integer = 1440
    Public Property MarginBottomTw As Integer = 1440
    Public Property MarginLeftTw As Integer = 1440
    Public Property MarginRightTw As Integer = 1440
    Public Property HeaderTw As Integer = 720
    Public Property FooterTw As Integer = 720

    Public ReadOnly Property ContentWidthTw As Integer
        Get
            Return WidthTw - MarginLeftTw - MarginRightTw
        End Get
    End Property
End Class

''' <summary>Seluruh isi dokumen yang dipakai generator.</summary>
Public Class DocModel
    Public Property Setup As New PageSetup
    Public Property Body As New List(Of DocBlock)
    Public Property Header As New List(Of DocBlock)
    Public Property Footer As New List(Of DocBlock)
    Public Property DefaultFont As String = "Segoe UI"
    Public Property DefaultSizeHalfPt As Integer = 16
    ''' <summary>nama file media -> byte isi file (ditulis ke folder output).</summary>
    Public Property Media As New Dictionary(Of String, Byte())
End Class
