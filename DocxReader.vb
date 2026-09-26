Option Strict On
Option Infer On

Imports System.IO
Imports System.IO.Compression
Imports System.Xml.Linq

' ============================================================================
'  DocxReader.vb  -- .docx (OOXML) -> DocModel
'  Dibaca langsung dari paket ZIP + XDocument, tanpa dependensi luar.
' ============================================================================

Public Module Ns
    Public ReadOnly WNs As XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    Public ReadOnly RNs As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    Public ReadOnly ANs As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
    Public ReadOnly WpNs As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    Public ReadOnly PkgNs As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
End Module

Public Class DocxReader

    Private _zip As ZipArchive
    Private _model As New DocModel

    ''' <summary>Baca seluruh isi .docx menjadi DocModel.</summary>
    Public Function Read(docxPath As String) As DocModel
        _model = New DocModel
        Using fs As New FileStream(docxPath, FileMode.Open, FileAccess.Read)
            _zip = New ZipArchive(fs, ZipArchiveMode.Read)

            ReadDefaults()

            Dim docRels = ReadRels("word/_rels/document.xml.rels")
            Dim doc = ReadXml("word/document.xml")
            Dim body = doc.Root.Element(Ns.WNs + "body")

            Dim sect = body.Element(Ns.WNs + "sectPr")
            ReadSectPr(sect)

            ' --- body ---
            For Each el In body.Elements()
                If el.Name = Ns.WNs + "p" Then
                    _model.Body.Add(DocBlock.FromPara(ParsePara(el, "word", docRels)))
                ElseIf el.Name = Ns.WNs + "tbl" Then
                    _model.Body.Add(DocBlock.FromTable(ParseTable(el, "word", docRels)))
                End If
            Next

            ' --- header & footer ---
            Dim hdrPart = RefTarget(sect, "headerReference", docRels)
            Dim ftrPart = RefTarget(sect, "footerReference", docRels)
            If hdrPart IsNot Nothing Then _model.Header = ReadHeaderFooter(hdrPart)
            If ftrPart IsNot Nothing Then _model.Footer = ReadHeaderFooter(ftrPart)
        End Using
        Return _model
    End Function

    ' ---------------------------------------------------------------- paket

    Private Function Entry(entryPath As String) As ZipArchiveEntry
        For Each e In _zip.Entries
            If String.Equals(e.FullName, entryPath, StringComparison.OrdinalIgnoreCase) Then Return e
        Next
        Return Nothing
    End Function

    Private Function ReadXml(xmlPath As String) As XDocument
        Dim e = Entry(xmlPath)
        If e Is Nothing Then Return Nothing
        Using s = e.Open()
            Return XDocument.Load(s)
        End Using
    End Function

    Private Function ReadBytes(binPath As String) As Byte()
        Dim e = Entry(binPath)
        If e Is Nothing Then Return Nothing
        Using s = e.Open(), ms As New MemoryStream()
            s.CopyTo(ms)
            Return ms.ToArray()
        End Using
    End Function

    ''' <summary>rId -> target relatif terhadap folder part.</summary>
    Private Function ReadRels(relsPath As String) As Dictionary(Of String, String)
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim x = ReadXml(relsPath)
        If x Is Nothing Then Return map
        For Each rel In x.Root.Elements(Ns.PkgNs + "Relationship")
            Dim id = CStr(rel.Attribute("Id"))
            Dim tgt = CStr(rel.Attribute("Target"))
            If id IsNot Nothing AndAlso tgt IsNot Nothing Then map(id) = tgt
        Next
        Return map
    End Function

    Private Function RefTarget(sect As XElement, refName As String,
                               rels As Dictionary(Of String, String)) As String
        If sect Is Nothing Then Return Nothing
        For Each el In sect.Elements(Ns.WNs + refName)
            Dim t = AttrS(el, Ns.WNs + "type")
            If t Is Nothing OrElse t = "default" Then
                Dim id = AttrS(el, Ns.RNs + "id")
                If id IsNot Nothing AndAlso rels.ContainsKey(id) Then Return "word/" & rels(id)
            End If
        Next
        Return Nothing
    End Function

    Private Function ReadHeaderFooter(partPath As String) As List(Of DocBlock)
        Dim res As New List(Of DocBlock)
        Dim x = ReadXml(partPath)
        If x Is Nothing Then Return res
        Dim dir = Path.GetDirectoryName(partPath).Replace("\", "/")
        Dim relsPath = dir & "/_rels/" & Path.GetFileName(partPath) & ".rels"
        Dim rels = ReadRels(relsPath)
        For Each el In x.Root.Elements()
            If el.Name = Ns.WNs + "p" Then
                res.Add(DocBlock.FromPara(ParsePara(el, dir, rels)))
            ElseIf el.Name = Ns.WNs + "tbl" Then
                res.Add(DocBlock.FromTable(ParseTable(el, dir, rels)))
            End If
        Next
        Return res
    End Function

    ' ------------------------------------------------------------ helper XML

    Private Shared Function AttrS(e As XElement, n As XName) As String
        If e Is Nothing Then Return Nothing
        Dim a = e.Attribute(n)
        If a Is Nothing Then Return Nothing
        Return a.Value
    End Function

    Private Shared Function AttrI(e As XElement, n As XName, dflt As Integer) As Integer
        Dim s = AttrS(e, n)
        Dim v As Integer
        If s IsNot Nothing AndAlso Integer.TryParse(s, v) Then Return v
        Return dflt
    End Function

    Private Shared Function AttrL(e As XElement, n As XName, dflt As Long) As Long
        Dim s = AttrS(e, n)
        Dim v As Long
        If s IsNot Nothing AndAlso Long.TryParse(s, v) Then Return v
        Return dflt
    End Function

    ''' <summary>Elemen on/off gaya Word: ada = true, kecuali w:val = 0/false/off.</summary>
    Private Shared Function OnOff(parent As XElement, name As String) As Boolean
        If parent Is Nothing Then Return False
        Dim e = parent.Element(Ns.WNs + name)
        If e Is Nothing Then Return False
        Dim v = AttrS(e, Ns.WNs + "val")
        If v Is Nothing Then Return True
        Return Not (v = "0" OrElse v = "false" OrElse v = "off")
    End Function

    Private Shared Function ValOf(parent As XElement, name As String) As String
        If parent Is Nothing Then Return Nothing
        Return AttrS(parent.Element(Ns.WNs + name), Ns.WNs + "val")
    End Function

    ' ------------------------------------------------------------- defaults

    Private Sub ReadDefaults()
        Dim st = ReadXml("word/styles.xml")
        If st Is Nothing Then Return
        Dim rPr = st.Root.Element(Ns.WNs + "docDefaults")?.
                          Element(Ns.WNs + "rPrDefault")?.Element(Ns.WNs + "rPr")
        If rPr Is Nothing Then Return
        Dim f = AttrS(rPr.Element(Ns.WNs + "rFonts"), Ns.WNs + "ascii")
        If Not String.IsNullOrEmpty(f) Then _model.DefaultFont = f
        Dim sz = AttrI(rPr.Element(Ns.WNs + "sz"), Ns.WNs + "val", 0)
        If sz > 0 Then _model.DefaultSizeHalfPt = sz
    End Sub

    Private Sub ReadSectPr(sect As XElement)
        If sect Is Nothing Then Return
        Dim s = _model.Setup
        Dim pgSz = sect.Element(Ns.WNs + "pgSz")
        s.WidthTw = AttrI(pgSz, Ns.WNs + "w", s.WidthTw)
        s.HeightTw = AttrI(pgSz, Ns.WNs + "h", s.HeightTw)
        Dim m = sect.Element(Ns.WNs + "pgMar")
        s.MarginTopTw = AttrI(m, Ns.WNs + "top", s.MarginTopTw)
        s.MarginBottomTw = AttrI(m, Ns.WNs + "bottom", s.MarginBottomTw)
        s.MarginLeftTw = AttrI(m, Ns.WNs + "left", s.MarginLeftTw)
        s.MarginRightTw = AttrI(m, Ns.WNs + "right", s.MarginRightTw)
        s.HeaderTw = AttrI(m, Ns.WNs + "header", s.HeaderTw)
        s.FooterTw = AttrI(m, Ns.WNs + "footer", s.FooterTw)
    End Sub

    ' ------------------------------------------------------------- paragraf

    Private Function ParsePara(pEl As XElement, partDir As String,
                               rels As Dictionary(Of String, String)) As DocPara
        Dim res As New DocPara
        Dim pPr = pEl.Element(Ns.WNs + "pPr")

        If pPr IsNot Nothing Then
            res.Align = ValOf(pPr, "jc")
            res.KeepNext = OnOff(pPr, "keepNext")

            Dim ind = pPr.Element(Ns.WNs + "ind")
            If ind IsNot Nothing Then
                res.IndentLeftTw = AttrI(ind, Ns.WNs + "left", 0)
                res.IndentRightTw = AttrI(ind, Ns.WNs + "right", 0)
                res.IndentFirstLineTw = AttrI(ind, Ns.WNs + "firstLine", 0)
                res.IndentHangingTw = AttrI(ind, Ns.WNs + "hanging", 0)
            End If

            Dim sp = pPr.Element(Ns.WNs + "spacing")
            If sp IsNot Nothing Then
                res.SpaceBeforeTw = AttrI(sp, Ns.WNs + "before", 0)
                res.SpaceAfterTw = AttrI(sp, Ns.WNs + "after", 0)
                res.LineTw = AttrI(sp, Ns.WNs + "line", 0)
                res.LineRule = AttrS(sp, Ns.WNs + "lineRule")
            End If

            Dim shd = pPr.Element(Ns.WNs + "shd")
            If shd IsNot Nothing Then
                Dim fill = AttrS(shd, Ns.WNs + "fill")
                If fill IsNot Nothing AndAlso fill <> "auto" Then res.ShadingHex = fill
            End If

            Dim markPr = pPr.Element(Ns.WNs + "rPr")
            If markPr IsNot Nothing Then
                res.MarkSizeHalfPt = AttrI(markPr.Element(Ns.WNs + "sz"), Ns.WNs + "val", 0)
            End If

            Dim tabs = pPr.Element(Ns.WNs + "tabs")
            If tabs IsNot Nothing Then
                For Each t In tabs.Elements(Ns.WNs + "tab")
                    res.TabStopsTw.Add(AttrI(t, Ns.WNs + "pos", 0))
                Next
            End If
        End If

        ParseRuns(pEl, res, partDir, rels)
        Return res
    End Function

    ''' <summary>Isi Runs; field PAGE/NUMPAGES dijadikan satu run placeholder.</summary>
    Private Sub ParseRuns(pEl As XElement, target As DocPara, partDir As String,
                          rels As Dictionary(Of String, String))
        Dim fieldState = 0          ' 0 = normal, 1 = baca instruksi, 2 = lewati hasil
        Dim instr As New Text.StringBuilder()

        For Each child In pEl.Elements()
            If child.Name = Ns.WNs + "hyperlink" Then
                For Each r In child.Elements(Ns.WNs + "r")
                    HandleRun(r, target, partDir, rels, fieldState, instr)
                Next
            ElseIf child.Name = Ns.WNs + "r" Then
                HandleRun(child, target, partDir, rels, fieldState, instr)
            End If
        Next
    End Sub

    Private Sub HandleRun(r As XElement, target As DocPara, partDir As String,
                          rels As Dictionary(Of String, String),
                          ByRef fieldState As Integer, instr As Text.StringBuilder)
        Dim rPr = r.Element(Ns.WNs + "rPr")

        For Each n In r.Elements()
            Select Case n.Name.LocalName
                Case "fldChar"
                    Select Case AttrS(n, Ns.WNs + "fldCharType")
                        Case "begin"
                            fieldState = 1
                            instr.Clear()
                        Case "separate"
                            fieldState = 2
                        Case "end"
                            Dim code = instr.ToString().Trim().ToUpperInvariant()
                            Dim ph = If(code.StartsWith("NUMPAGES"), "{{PAGES}}",
                                     If(code.StartsWith("PAGE"), "{{PAGE}}", ""))
                            If ph <> "" Then
                                Dim fr = NewRun(rPr)
                                fr.Kind = RunKind.Field
                                fr.Text = ph
                                target.Runs.Add(fr)
                            End If
                            fieldState = 0
                    End Select

                Case "instrText"
                    If fieldState = 1 Then instr.Append(n.Value)

                Case "t"
                    If fieldState = 0 Then
                        Dim tr = NewRun(rPr)
                        tr.Kind = RunKind.Text
                        tr.Text = n.Value
                        target.Runs.Add(tr)
                    End If

                Case "tab"
                    If fieldState = 0 Then
                        Dim tr = NewRun(rPr)
                        tr.Kind = RunKind.Tab
                        target.Runs.Add(tr)
                    End If

                Case "br"
                    If fieldState = 0 Then
                        Dim tr = NewRun(rPr)
                        tr.Kind = If(AttrS(n, Ns.WNs + "type") = "page",
                                     RunKind.PageBreak, RunKind.LineBreak)
                        target.Runs.Add(tr)
                    End If

                Case "drawing", "pict"
                    If fieldState = 0 Then
                        Dim pic = ParsePicture(n, partDir, rels)
                        If pic IsNot Nothing Then target.Runs.Add(pic)
                    End If
            End Select
        Next
    End Sub

    Private Function NewRun(rPr As XElement) As DocRun
        Dim d As New DocRun
        If rPr Is Nothing Then Return d
        d.Bold = OnOff(rPr, "b")
        d.Italic = OnOff(rPr, "i")
        d.Underline = (ValOf(rPr, "u") IsNot Nothing AndAlso ValOf(rPr, "u") <> "none")
        Dim va = ValOf(rPr, "vertAlign")
        d.Superscript = (va = "superscript")
        d.Subscript = (va = "subscript")
        d.FontName = AttrS(rPr.Element(Ns.WNs + "rFonts"), Ns.WNs + "ascii")
        d.SizeHalfPt = AttrI(rPr.Element(Ns.WNs + "sz"), Ns.WNs + "val", 0)
        Dim c = ValOf(rPr, "color")
        If c IsNot Nothing AndAlso c <> "auto" Then d.ColorHex = c
        Dim hl = ValOf(rPr, "highlight")
        If hl IsNot Nothing AndAlso hl <> "none" Then d.HighlightName = hl
        Return d
    End Function

    Private Function ParsePicture(n As XElement, partDir As String,
                                  rels As Dictionary(Of String, String)) As DocRun
        Dim ext = n.Descendants(Ns.WpNs + "extent").FirstOrDefault()
        Dim blip = n.Descendants(Ns.ANs + "blip").FirstOrDefault()
        If blip Is Nothing Then Return Nothing
        Dim embed = AttrS(blip, Ns.RNs + "embed")
        If embed Is Nothing OrElse Not rels.ContainsKey(embed) Then Return Nothing

        Dim tgt = rels(embed).Replace("\", "/")
        Dim full = NormalizePath(partDir & "/" & tgt)
        Dim bytes = ReadBytes(full)
        If bytes Is Nothing Then Return Nothing

        Dim name = "media/" & Path.GetFileName(full)
        _model.Media(name) = bytes

        ' wp:extent memakai atribut tanpa namespace: cx / cy (EMU)
        Return New DocRun With {
            .Kind = RunKind.Picture,
            .ImageFile = name,
            .WidthEmu = AttrL(ext, XName.Get("cx"), 0),
            .HeightEmu = AttrL(ext, XName.Get("cy"), 0)
        }
    End Function

    ''' <summary>Rapikan "word/../word/media/x.png" jadi "word/media/x.png".</summary>
    Private Shared Function NormalizePath(raw As String) As String
        Dim parts As New List(Of String)
        For Each seg In raw.Split("/"c)
            If seg = "." OrElse seg = "" Then Continue For
            If seg = ".." Then
                If parts.Count > 0 Then parts.RemoveAt(parts.Count - 1)
            Else
                parts.Add(seg)
            End If
        Next
        Return String.Join("/", parts)
    End Function

    ' ---------------------------------------------------------------- tabel

    Private Function ParseTable(tbl As XElement, partDir As String,
                                rels As Dictionary(Of String, String)) As DocTable
        Dim res As New DocTable
        Dim tblPr = tbl.Element(Ns.WNs + "tblPr")

        If tblPr IsNot Nothing Then
            res.IndentTw = AttrI(tblPr.Element(Ns.WNs + "tblInd"), Ns.WNs + "w", 0)
            res.Borders = ParseBorders(tblPr.Element(Ns.WNs + "tblBorders"),
                                       {"top", "left", "bottom", "right", "insideH", "insideV"})
            res.CellMarginsTw = ParseMargins(tblPr.Element(Ns.WNs + "tblCellMar"))
        End If

        Dim grid = tbl.Element(Ns.WNs + "tblGrid")
        If grid IsNot Nothing Then
            For Each gc In grid.Elements(Ns.WNs + "gridCol")
                res.GridTw.Add(AttrI(gc, Ns.WNs + "w", 0))
            Next
        End If

        For Each tr In tbl.Elements(Ns.WNs + "tr")
            Dim row As New DocRow
            Dim trPr = tr.Element(Ns.WNs + "trPr")
            If trPr IsNot Nothing Then
                Dim h = trPr.Element(Ns.WNs + "trHeight")
                row.HeightTw = AttrI(h, Ns.WNs + "val", 0)
                row.HeightRule = AttrS(h, Ns.WNs + "hRule")
                row.CantSplit = OnOff(trPr, "cantSplit")
            End If

            For Each tc In tr.Elements(Ns.WNs + "tc")
                Dim cell As New DocCell
                Dim tcPr = tc.Element(Ns.WNs + "tcPr")
                If tcPr IsNot Nothing Then
                    cell.WidthTw = AttrI(tcPr.Element(Ns.WNs + "tcW"), Ns.WNs + "w", 0)
                    cell.GridSpan = AttrI(tcPr.Element(Ns.WNs + "gridSpan"), Ns.WNs + "val", 1)
                    Dim vm = tcPr.Element(Ns.WNs + "vMerge")
                    If vm IsNot Nothing Then
                        cell.VMerge = If(AttrS(vm, Ns.WNs + "val"), "continue")
                    End If
                    cell.VAlign = ValOf(tcPr, "vAlign")
                    Dim shd = tcPr.Element(Ns.WNs + "shd")
                    If shd IsNot Nothing Then
                        Dim fill = AttrS(shd, Ns.WNs + "fill")
                        If fill IsNot Nothing AndAlso fill <> "auto" Then cell.ShadingHex = fill
                    End If
                    cell.Borders = ParseBorders(tcPr.Element(Ns.WNs + "tcBorders"),
                                                {"top", "left", "bottom", "right"})
                    cell.MarginsTw = ParseMargins(tcPr.Element(Ns.WNs + "tcMar"))
                End If

                For Each el In tc.Elements()
                    If el.Name = Ns.WNs + "p" Then
                        cell.Blocks.Add(DocBlock.FromPara(ParsePara(el, partDir, rels)))
                    ElseIf el.Name = Ns.WNs + "tbl" Then
                        cell.Blocks.Add(DocBlock.FromTable(ParseTable(el, partDir, rels)))
                    End If
                Next
                row.Cells.Add(cell)
            Next
            res.Rows.Add(row)
        Next

        ResolveVMerge(res)
        Return res
    End Function

    Private Shared Function ParseBorders(el As XElement, names As String()) As Dictionary(Of String, DocBorder)
        Dim map As New Dictionary(Of String, DocBorder)(StringComparer.OrdinalIgnoreCase)
        If el Is Nothing Then Return map
        For Each n In names
            Dim b = el.Element(Ns.WNs + n)
            If b Is Nothing Then Continue For
            map(n) = New DocBorder With {
                .Val = If(AttrS(b, Ns.WNs + "val"), "none"),
                .SizeEighthPt = AttrI(b, Ns.WNs + "sz", 0),
                .ColorHex = If(AttrS(b, Ns.WNs + "color"), "auto")
            }
        Next
        Return map
    End Function

    ''' <summary>tblCellMar / tcMar -> {top, left, bottom, right} dalam twip.</summary>
    Private Shared Function ParseMargins(el As XElement) As Integer()
        If el Is Nothing Then Return Nothing
        Dim res = New Integer() {-1, -1, -1, -1}
        Dim names = {"top", "left", "bottom", "right"}
        For i = 0 To 3
            Dim m = el.Element(Ns.WNs + names(i))
            If m IsNot Nothing Then res(i) = AttrI(m, Ns.WNs + "w", 0)
        Next
        Return res
    End Function

    ''' <summary>
    ''' vMerge restart/continue -> RowSpan pada sel pertama, sisanya Skip.
    ''' Pencocokan memakai posisi kolom grid (bukan indeks sel), karena sel
    ''' ber-gridSpan menggeser indeks antar baris.
    ''' </summary>
    Private Shared Sub ResolveVMerge(t As DocTable)
        ' posisi kolom grid awal tiap sel, per baris
        Dim gridPos As New List(Of Dictionary(Of Integer, DocCell))
        For Each row In t.Rows
            Dim map As New Dictionary(Of Integer, DocCell)
            Dim col = 0
            For Each c In row.Cells
                map(col) = c
                col += Math.Max(1, c.GridSpan)
            Next
            gridPos.Add(map)
        Next

        For ri = 0 To t.Rows.Count - 1
            For Each kv In gridPos(ri)
                Dim c = kv.Value
                If c.VMerge Is Nothing OrElse c.VMerge <> "restart" Then Continue For
                Dim span = 1
                Dim rj = ri + 1
                While rj < t.Rows.Count
                    Dim below As DocCell = Nothing
                    If Not gridPos(rj).TryGetValue(kv.Key, below) Then Exit While
                    If below.VMerge Is Nothing OrElse below.VMerge <> "continue" Then Exit While
                    below.Skip = True
                    span += 1
                    rj += 1
                End While
                c.RowSpan = span
            Next
        Next

        ' sel "continue" tanpa restart di atasnya (jarang) -> tetap ditulis biasa
        For Each row In t.Rows
            For Each c In row.Cells
                If c.VMerge = "continue" AndAlso Not c.Skip Then c.VMerge = Nothing
            Next
        Next
    End Sub

End Class
