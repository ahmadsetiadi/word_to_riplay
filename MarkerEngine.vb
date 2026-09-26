Option Strict On
Option Infer On

Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

' ============================================================================
'  MarkerEngine.vb -- jalankan semua penanda <<...>> di atas DocModel,
'  SEBELUM HTML ditulis. Hasilnya HTML tanpa <<...>> tersisa.
'
'  Urutan kerja tiap daftar blok:
'    1. <<if EXPR>> … <<end if>>   blok tampil/sembunyi (boleh bersarang)
'    2. <<row-if EXPR>>            baris tabel tampil/sembunyi (tanpa penutup)
'    3. <<arr.field>>              baris tabel berulang, satu baris per item
'    4. <<apa pun>>                diganti nilainya
'    5. penomoran ulang            1. 2. 3. / a. b. c. dirapikan
'
'  Bekerja di level model (paragraf, tabel, baris, sel) — bukan regex di HTML —
'  jadi cakupan blok dan baris tabel bisa ditentukan dengan tepat.
' ============================================================================

Public Class MarkerEngine

    Private ReadOnly _data As Dictionary(Of String, Object)
    Private ReadOnly _log As Action(Of String)
    Private ReadOnly _ev As ExprEngine

    Public Property Replaced As Integer = 0
    Public Property Unknown As New List(Of String)
    Public Property BlocksKept As Integer = 0
    Public Property BlocksDropped As Integer = 0
    Public Property RowsDropped As Integer = 0
    Public Property RowsRepeated As Integer = 0
    Public Property Renumbered As Integer = 0

    Public Sub New(data As Dictionary(Of String, Object), log As Action(Of String))
        _data = data
        _log = log
        _ev = New ExprEngine(data, log)
    End Sub

    Private Sub Say(m As String)
        If _log IsNot Nothing Then _log(m)
    End Sub

    ' penanda: <<...>> tanpa << atau >> di dalamnya
    Private Shared ReadOnly MarkRx As New Regex("<<(?<body>(?:(?!<<|>>).)*)>>", RegexOptions.Compiled)
    Private Shared ReadOnly IfSpaceRx As New Regex("^\s*if\s+(?<e>.+?)\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private Shared ReadOnly IfCallRx As New Regex("^\s*if\s*\((?<e>.*)\)\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline)
    Private Shared ReadOnly EndIfRx As New Regex("^\s*(?:end\s*if|endif|/\s*if)\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private Shared ReadOnly RowIfRx As New Regex("^\s*row-if\s+(?<e>.+?)\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private Shared ReadOnly PageBreakRx As New Regex("^\s*page\s*[-_ ]?\s*break\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private Shared ReadOnly FieldRx As New Regex("^\s*(?<arr>[A-Za-z_$][\w$]*)\.(?<fld>[A-Za-z_$][\w$]*)\s*$", RegexOptions.Compiled)

    ''' <summary>Jalankan seluruh penanda pada model.</summary>
    Public Sub Apply(m As DocModel)
        ProcessBlocks(m.Body)
        ProcessBlocks(m.Header)
        ProcessBlocks(m.Footer)
        Renumber(m.Body, New Dictionary(Of String, Integer), "body")
    End Sub

    ' ==================================================== 1. blok <<if>>…<<end if>>

    Private Sub ProcessBlocks(blocks As List(Of DocBlock))
        ResolveIfBlocks(blocks)
        For Each b In blocks
            If b.IsTable Then
                ProcessTable(b.Table)
            Else
                SubstitutePara(b.Para)
            End If
        Next
    End Sub

    ''' <summary>Paragraf yang isinya hanya &lt;&lt;if …&gt;&gt; membuka blok; &lt;&lt;end if&gt;&gt; menutupnya.</summary>
    Private Sub ResolveIfBlocks(blocks As List(Of DocBlock))
        Dim i = 0
        While i < blocks.Count
            Dim expr = LoneIf(blocks(i))
            If expr Is Nothing Then
                i += 1
                Continue While
            End If

            ' cari <<end if>> pasangannya (hitung yang bersarang)
            Dim depth = 1
            Dim j = i + 1
            While j < blocks.Count
                If LoneIf(blocks(j)) IsNot Nothing Then
                    depth += 1
                ElseIf IsLoneEndIf(blocks(j)) Then
                    depth -= 1
                    If depth = 0 Then Exit While
                End If
                j += 1
            End While

            If j >= blocks.Count Then
                Say("    ! <<if " & expr & ">> tidak punya <<end if>> — blok diabaikan (isi tetap tampil)")
                blocks.RemoveAt(i)
                Continue While
            End If

            Dim keep = _ev.EvalBool(expr)
            If keep Then
                BlocksKept += 1
                blocks.RemoveAt(j)       ' buang <<end if>>
                blocks.RemoveAt(i)       ' buang <<if>>
                ' isi tetap, lanjut periksa dari posisi yang sama (blok bersarang)
            Else
                BlocksDropped += 1
                blocks.RemoveRange(i, j - i + 1)
            End If
            Say("    · <<if " & Trunc(expr) & ">> -> " & If(keep, "TAMPIL", "disembunyikan"))
        End While
    End Sub

    ''' <summary>
    ''' Kondisi blok, kalau paragraf ini HANYA berisi penanda pembuka blok.
    ''' Diterima: &lt;&lt;if EXPR&gt;&gt; (pakai spasi) dan &lt;&lt;if(EXPR)&gt;&gt; dengan SATU argumen.
    ''' &lt;&lt;if(kondisi, a, b)&gt;&gt; (2-3 argumen) adalah NILAI, bukan blok — dibedakan
    ''' dari jumlah koma di tingkat teratas.
    ''' </summary>
    Private Shared Function LoneIf(b As DocBlock) As String
        If Not b.IsPara Then Return Nothing
        Dim t = b.Para.PlainText.Trim()
        Dim m = MarkRx.Match(t)
        If Not m.Success OrElse m.Value <> t Then Return Nothing
        Dim body = ExprEngine.Normalize(m.Groups("body").Value)

        Dim asCall = IfCallRx.Match(body)
        If asCall.Success Then
            If TopLevelCommas(asCall.Groups("e").Value) = 0 Then Return asCall.Groups("e").Value
            Return Nothing
        End If
        Dim asSpace = IfSpaceRx.Match(body)
        If asSpace.Success Then Return asSpace.Groups("e").Value
        Return Nothing
    End Function

    ''' <summary>Hitung koma pada kedalaman kurung 0 dan di luar kutip.</summary>
    Private Shared Function TopLevelCommas(s As String) As Integer
        Dim depth = 0
        Dim n = 0
        Dim q As Char = ChrW(0)
        For Each ch In s
            If q <> ChrW(0) Then
                If ch = q Then q = ChrW(0)
            ElseIf ch = "'"c OrElse ch = ChrW(34) Then
                q = ch
            ElseIf ch = "("c Then
                depth += 1
            ElseIf ch = ")"c Then
                depth -= 1
            ElseIf ch = ","c AndAlso depth = 0 Then
                n += 1
            End If
        Next
        Return n
    End Function

    Private Function LoneMarker(b As DocBlock, rx As Regex, grp As String) As String
        If Not b.IsPara Then Return Nothing
        Dim t = b.Para.PlainText.Trim()
        Dim m = MarkRx.Match(t)
        If Not m.Success OrElse m.Value <> t Then Return Nothing
        Dim inner = rx.Match(m.Groups("body").Value)
        If Not inner.Success Then Return Nothing
        Return inner.Groups(grp).Value
    End Function

    Private Function IsLoneEndIf(b As DocBlock) As Boolean
        If Not b.IsPara Then Return False
        Dim t = b.Para.PlainText.Trim()
        Dim m = MarkRx.Match(t)
        If Not m.Success OrElse m.Value <> t Then Return False
        Return EndIfRx.IsMatch(m.Groups("body").Value)
    End Function

    ' ==================================================== 2-3. tabel

    Private Sub ProcessTable(t As DocTable)
        ' --- row-if: buang baris atau hapus penandanya ---
        Dim ri = 0
        While ri < t.Rows.Count
            Dim cond = FindRowIf(t.Rows(ri))
            If cond IsNot Nothing Then
                Dim keep = _ev.EvalBool(cond)
                Say("    · <<row-if " & Trunc(cond) & ">> -> " & If(keep, "TAMPIL", "baris dibuang"))
                If Not keep Then
                    t.Rows.RemoveAt(ri)
                    RowsDropped += 1
                    Continue While
                End If
                StripRowIf(t.Rows(ri))
            End If
            ri += 1
        End While

        ' --- baris berulang <<arr.field>> ---
        ExpandRepeatRows(t)

        ' --- isi sel ---
        For Each row In t.Rows
            For Each c In row.Cells
                ProcessBlocks(c.Blocks)
            Next
        Next
    End Sub

    Private Function FindRowIf(row As DocRow) As String
        For Each c In row.Cells
            For Each b In c.Blocks
                If Not b.IsPara Then Continue For
                For Each m As Match In MarkRx.Matches(b.Para.PlainText)
                    Dim r = RowIfRx.Match(m.Groups("body").Value)
                    If r.Success Then Return r.Groups("e").Value
                Next
            Next
        Next
        Return Nothing
    End Function

    Private Sub StripRowIf(row As DocRow)
        For Each c In row.Cells
            For Each b In c.Blocks
                If b.IsPara Then RemoveMarkers(b.Para, RowIfRx)
            Next
        Next
    End Sub

    ''' <summary>Baris yang memuat &lt;&lt;arr.field&gt;&gt; digandakan satu kali per item array.</summary>
    Private Sub ExpandRepeatRows(t As DocTable)
        Dim ri = 0
        While ri < t.Rows.Count
            Dim arrName = FindArrayName(t.Rows(ri))
            If arrName Is Nothing Then
                ri += 1
                Continue While
            End If
            Dim items = TryCast(GetData(arrName), List(Of Object))
            If items Is Nothing OrElse items.Count = 0 Then
                Say("    ! <<" & arrName & ".*>> : data '" & arrName & "' kosong / bukan array — baris dibiarkan")
                ri += 1
                Continue While
            End If

            Dim template = t.Rows(ri)
            t.Rows.RemoveAt(ri)
            Dim n = 0
            For Each it In items
                Dim clone = CloneRow(template)
                Dim scope = TryCast(it, Dictionary(Of String, Object))
                FillRow(clone, arrName, scope, n + 1)
                t.Rows.Insert(ri + n, clone)
                n += 1
            Next
            RowsRepeated += n
            Say("    · <<" & arrName & ".*>> -> " & n.ToString(CultureInfo.InvariantCulture) & " baris")
            ri += n
        End While
    End Sub

    Private Function FindArrayName(row As DocRow) As String
        For Each c In row.Cells
            For Each b In c.Blocks
                If Not b.IsPara Then Continue For
                For Each m As Match In MarkRx.Matches(b.Para.PlainText)
                    Dim f = FieldRx.Match(m.Groups("body").Value)
                    If f.Success Then
                        Dim nm = f.Groups("arr").Value
                        If TypeOf GetData(nm) Is List(Of Object) Then Return nm
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function

    ' --- salinan dalam (deep clone) untuk baris berulang ---

    Private Shared Function CloneRow(r As DocRow) As DocRow
        Dim n As New DocRow With {
            .HeightTw = r.HeightTw, .HeightRule = r.HeightRule, .CantSplit = r.CantSplit}
        For Each c In r.Cells
            n.Cells.Add(CloneCell(c))
        Next
        Return n
    End Function

    Private Shared Function CloneCell(c As DocCell) As DocCell
        Dim n As New DocCell With {
            .WidthTw = c.WidthTw, .GridSpan = c.GridSpan, .VMerge = c.VMerge,
            .VAlign = c.VAlign, .ShadingHex = c.ShadingHex,
            .MarginsTw = If(c.MarginsTw Is Nothing, Nothing, CType(c.MarginsTw.Clone(), Integer())),
            .RowSpan = c.RowSpan, .Skip = c.Skip}
        For Each kv In c.Borders
            n.Borders(kv.Key) = New DocBorder With {
                .Val = kv.Value.Val, .SizeEighthPt = kv.Value.SizeEighthPt, .ColorHex = kv.Value.ColorHex}
        Next
        For Each b In c.Blocks
            n.Blocks.Add(CloneBlock(b))
        Next
        Return n
    End Function

    Private Shared Function CloneBlock(b As DocBlock) As DocBlock
        If b.IsPara Then Return DocBlock.FromPara(ClonePara(b.Para))
        Return DocBlock.FromTable(CloneTable(b.Table))
    End Function

    Private Shared Function CloneTable(t As DocTable) As DocTable
        Dim n As New DocTable With {
            .IndentTw = t.IndentTw,
            .CellMarginsTw = If(t.CellMarginsTw Is Nothing, Nothing, CType(t.CellMarginsTw.Clone(), Integer()))}
        n.GridTw.AddRange(t.GridTw)
        For Each kv In t.Borders
            n.Borders(kv.Key) = New DocBorder With {
                .Val = kv.Value.Val, .SizeEighthPt = kv.Value.SizeEighthPt, .ColorHex = kv.Value.ColorHex}
        Next
        For Each r In t.Rows
            n.Rows.Add(CloneRow(r))
        Next
        Return n
    End Function

    Private Shared Function ClonePara(p As DocPara) As DocPara
        Dim n As New DocPara With {
            .Align = p.Align, .IndentLeftTw = p.IndentLeftTw, .IndentRightTw = p.IndentRightTw,
            .IndentFirstLineTw = p.IndentFirstLineTw, .IndentHangingTw = p.IndentHangingTw,
            .SpaceBeforeTw = p.SpaceBeforeTw, .SpaceAfterTw = p.SpaceAfterTw,
            .LineTw = p.LineTw, .LineRule = p.LineRule, .ShadingHex = p.ShadingHex,
            .KeepNext = p.KeepNext, .MarkSizeHalfPt = p.MarkSizeHalfPt}
        n.TabStopsTw.AddRange(p.TabStopsTw)
        For Each r In p.Runs
            n.Runs.Add(New DocRun With {
                .Kind = r.Kind, .Text = r.Text, .Bold = r.Bold, .Italic = r.Italic,
                .Underline = r.Underline, .Superscript = r.Superscript, .Subscript = r.Subscript,
                .FontName = r.FontName, .SizeHalfPt = r.SizeHalfPt, .ColorHex = r.ColorHex,
                .HighlightName = r.HighlightName, .ImageFile = r.ImageFile,
                .WidthEmu = r.WidthEmu, .HeightEmu = r.HeightEmu})
        Next
        Return n
    End Function

    Private Function GetData(name As String) As Object
        If _data.ContainsKey(name) Then Return _data(name)
        Dim hit = _data.Keys.FirstOrDefault(Function(k) String.Equals(k, name, StringComparison.OrdinalIgnoreCase))
        Return If(hit Is Nothing, Nothing, _data(hit))
    End Function

    Private Sub FillRow(row As DocRow, arrName As String,
                        item As Dictionary(Of String, Object), no As Integer)
        For Each c In row.Cells
            For Each b In c.Blocks
                If Not b.IsPara Then Continue For
                ReplaceInPara(b.Para,
                    Function(body)
                        Dim f = FieldRx.Match(body)
                        If Not f.Success OrElse
                           Not String.Equals(f.Groups("arr").Value, arrName, StringComparison.OrdinalIgnoreCase) Then
                            Return Nothing
                        End If
                        Dim fld = f.Groups("fld").Value
                        If String.Equals(fld, "no", StringComparison.OrdinalIgnoreCase) Then
                            Return no.ToString(CultureInfo.InvariantCulture)
                        End If
                        If item Is Nothing Then Return ""
                        Dim hit = item.Keys.FirstOrDefault(Function(k) String.Equals(k, fld, StringComparison.OrdinalIgnoreCase))
                        Return If(hit Is Nothing, "", ExprEngine.ToText(item(hit)))
                    End Function)
            Next
        Next
    End Sub

    ' ==================================================== 4. penggantian nilai

    Private Sub SubstitutePara(p As DocPara)
        ReplaceInPara(p, Function(body)
                             If PageBreakRx.IsMatch(body) Then Return Nothing   ' biarkan
                             Dim bn = body.Trim().ToLowerInvariant()
                             If bn = "pageno" OrElse bn = "page" Then Return "{{PAGE}}"
                             If bn = "totalpage" OrElse bn = "totalpages" Then Return "{{PAGES}}"
                             If EndIfRx.IsMatch(body) Then Return ""
                             If RowIfRx.IsMatch(body) Then Return ""
                             Dim v = _ev.EvalValue(body)
                             If v Is Nothing AndAlso Not _ev.KnowsRoot(body) Then
                                 If Not Unknown.Contains(body) Then Unknown.Add(body)
                                 Return Nothing                                  ' biarkan merah
                             End If
                             Replaced += 1
                             Return ExprEngine.ToText(v)
                         End Function)
    End Sub

    ''' <summary>
    ''' Ganti penanda di seluruh paragraf. resolve(isi) mengembalikan teks pengganti,
    ''' atau Nothing kalau penanda harus dibiarkan apa adanya.
    ''' Penanda yang terbentang di beberapa run ditangani: nilai masuk ke run pertama,
    ''' potongan di run berikutnya dikosongkan.
    ''' </summary>
    Private Shared Sub ReplaceInPara(p As DocPara, resolve As Func(Of String, String))
        If p.Runs.Count = 0 Then Return

        ' peta posisi karakter -> run
        Dim sb As New StringBuilder()
        Dim span As New List(Of Integer())     ' {runIndex, start, end}
        For i = 0 To p.Runs.Count - 1
            Dim r = p.Runs(i)
            If r.Kind <> RunKind.Text AndAlso r.Kind <> RunKind.Field Then Continue For
            Dim s = sb.Length
            sb.Append(r.Text)
            span.Add(New Integer() {i, s, sb.Length})
        Next
        Dim full = sb.ToString()
        If full.IndexOf("<<", StringComparison.Ordinal) < 0 Then Return

        Dim ms = MarkRx.Matches(full)
        If ms.Count = 0 Then Return

        ' proses dari belakang supaya indeks tidak bergeser
        Dim texts As New Dictionary(Of Integer, String)
        For Each sp In span
            texts(sp(0)) = p.Runs(sp(0)).Text
        Next

        For k = ms.Count - 1 To 0 Step -1
            Dim m = ms(k)
            Dim rep = resolve(m.Groups("body").Value)
            If rep Is Nothing Then Continue For

            Dim first = True
            For Each sp In span
                Dim ri = sp(0), a = sp(1), b = sp(2)
                If b <= m.Index OrElse a >= m.Index + m.Length Then Continue For
                Dim lo = Math.Max(a, m.Index) - a
                Dim hi = Math.Min(b, m.Index + m.Length) - a
                Dim cur = texts(ri)
                texts(ri) = cur.Substring(0, lo) & If(first, rep, "") & cur.Substring(hi)
                first = False
            Next
            ' indeks span perlu dihitung ulang kalau masih ada penanda lain di run yang sama
            Dim sb2 As New StringBuilder()
            Dim span2 As New List(Of Integer())
            For Each sp In span
                Dim s2 = sb2.Length
                sb2.Append(texts(sp(0)))
                span2.Add(New Integer() {sp(0), s2, sb2.Length})
            Next
            span = span2
            full = sb2.ToString()
            ms = MarkRx.Matches(full)
            k = Math.Min(k, ms.Count)
        Next

        For Each kv In texts
            p.Runs(kv.Key).Text = kv.Value
        Next
    End Sub

    Private Shared Sub RemoveMarkers(p As DocPara, rx As Regex)
        ReplaceInPara(p, Function(body) If(rx.IsMatch(body), "", Nothing))
    End Sub

    ' ==================================================== 5. penomoran ulang

    Private Shared ReadOnly NumLabelRx As New Regex("^\s*(?<n>\d+)\.\s*$", RegexOptions.Compiled)
    Private Shared ReadOnly AlphaLabelRx As New Regex("^\s*(?<a>[a-z])\.\s*$", RegexOptions.Compiled)
    Private Shared ReadOnly RomanLabelRx As New Regex("^\s*(?<r>i{1,3}v?|iv|v|vi{1,3}|ix|x)\.\s*$", RegexOptions.Compiled)

    ''' <summary>
    ''' Rapikan nomor 1. 2. 3. / a. b. c. / i. ii. iii. setelah ada blok yang dibuang.
    ''' Level ditentukan dari indent kiri paragraf; label yang tidak berbentuk
    ''' nomor dibiarkan apa adanya.
    ''' </summary>
    Private _tableSeq As Integer = 0

    Private Sub Renumber(blocks As List(Of DocBlock),
                         counters As Dictionary(Of String, Integer),
                         scope As String)
        For Each b In blocks
            If b.IsTable Then
                ' Nomor berlanjut ke bawah dalam KOLOM yang sama pada satu tabel
                ' (tabel manfaat: 1./2./3. di kolom 1 beberapa baris),
                ' tapi tiap kolom berdiri sendiri (tata letak 2 kolom halaman 3-6).
                _tableSeq += 1
                Dim tid = scope & "/t" & _tableSeq.ToString(CultureInfo.InvariantCulture)
                For Each row In b.Table.Rows
                    Dim ci = 0
                    For Each c In row.Cells
                        Renumber(c.Blocks, counters, tid & "c" & ci.ToString(CultureInfo.InvariantCulture))
                        ci += Math.Max(1, c.GridSpan)
                    Next
                Next
                Continue For
            End If

            Dim p = b.Para
            If p.IndentHangingTw <= 0 OrElse p.Runs.Count = 0 Then Continue For

            ' label = run-run sebelum tab pertama
            Dim tabAt = -1
            For i = 0 To p.Runs.Count - 1
                If p.Runs(i).Kind = RunKind.Tab Then
                    tabAt = i
                    Exit For
                End If
            Next
            If tabAt <= 0 Then Continue For

            Dim label As New StringBuilder()
            For i = 0 To tabAt - 1
                If p.Runs(i).Kind = RunKind.Text Then label.Append(p.Runs(i).Text)
            Next
            Dim lbl = label.ToString()

            Dim kind = 0
            If NumLabelRx.IsMatch(lbl) Then
                kind = 1
            ElseIf AlphaLabelRx.IsMatch(lbl) Then
                kind = 2
            ElseIf RomanLabelRx.IsMatch(lbl) Then
                kind = 3
            Else
                Continue For
            End If

            Dim lvl = p.IndentLeftTw
            Dim key = scope & "|" & lvl.ToString(CultureInfo.InvariantCulture)
            If Not counters.ContainsKey(key) Then counters(key) = 0
            ' level yang lebih dalam pada scope yang sama di-reset saat level ini maju
            For Each k In counters.Keys.ToList()
                If Not k.StartsWith(scope & "|", StringComparison.Ordinal) Then Continue For
                Dim other As Integer
                If Integer.TryParse(k.Substring(scope.Length + 1), other) AndAlso other > lvl Then
                    counters(k) = 0
                End If
            Next
            counters(key) += 1
            Dim newLbl = MakeLabel(kind, counters(key))

            If newLbl <> lbl.Trim() Then
                Renumbered += 1
                ' tulis label baru ke run teks pertama, kosongkan sisanya
                Dim written = False
                For i = 0 To tabAt - 1
                    If p.Runs(i).Kind <> RunKind.Text Then Continue For
                    p.Runs(i).Text = If(written, "", newLbl)
                    written = True
                Next
            End If
        Next
    End Sub

    Private Shared Function MakeLabel(kind As Integer, n As Integer) As String
        Select Case kind
            Case 1 : Return n.ToString(CultureInfo.InvariantCulture) & "."
            Case 2 : Return Chr(Asc("a"c) + ((n - 1) Mod 26)) & "."
            Case Else : Return Roman(n).ToLowerInvariant() & "."
        End Select
    End Function

    Private Shared Function Roman(n As Integer) As String
        Dim vals = {10, 9, 5, 4, 1}
        Dim sym = {"X", "IX", "V", "IV", "I"}
        Dim sb As New StringBuilder()
        For i = 0 To vals.Length - 1
            While n >= vals(i)
                sb.Append(sym(i))
                n -= vals(i)
            End While
        Next
        Return sb.ToString()
    End Function

    Private Shared Function Trunc(s As String) As String
        Return If(s.Length > 55, s.Substring(0, 55) & "…", s)
    End Function

End Class
