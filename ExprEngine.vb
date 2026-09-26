Option Strict On
Option Infer On

Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions

' ============================================================================
'  ExprEngine.vb -- evaluasi isi penanda <<...>> menjadi NILAI (bukan cuma benar/salah)
'
'  Didukung:
'    nilai      : 'teks'  "teks"  123  4.5  true  false  null
'    variabel   : insuredname   policy.cur   riders
'    aritmatika : + - * /            (120 - insuredentryage)
'    banding    : = == != <> < <= > >=
'    logika     : and && or || not !
'    fungsi     : if(kondisi, nilaiBenar, nilaiSalah)
'                 format(nilai, 'dd/mm/yyyy')  /  format(nilai, '#,##0.00')
'    method     : .contains(x) .includes(x) .length .count
'                 .startsWith(x) .endsWith(x) .toLowerCase() .toUpperCase() .trim()
'
'  Benar/salah: boolean apa adanya; angka <> 0; teks tidak kosong;
'               array/objek tidak kosong; null / tidak ada = salah.
'
'  Kesalahan penulisan yang sering terjadi diperbaiki otomatis oleh Repair()
'  dan SELALU dicatat di log -- tidak ada yang diperbaiki diam-diam.
' ============================================================================

Public Class ExprEngine

    Private ReadOnly _data As Dictionary(Of String, Object)
    Private ReadOnly _log As Action(Of String)
    Private _toks As List(Of String)
    Private _pos As Integer

    Public Sub New(data As Dictionary(Of String, Object), log As Action(Of String))
        _data = data
        _log = log
    End Sub

    Private Sub Say(msg As String)
        If _log IsNot Nothing Then _log(msg)
    End Sub

    ' =====================================================  perbaikan otomatis

    ''' <summary>Samakan kutip keriting / tanda minus panjang jadi bentuk baku.</summary>
    Public Shared Function Normalize(expr As String) As String
        Dim s = Regex.Replace(expr, "[‘’‚‛]", "'")
        s = Regex.Replace(s, "[“”„‟]", """")
        s = Regex.Replace(s, "[–—−]", "-")
        s = s.Replace(ChrW(160), " "c)                    ' &nbsp; -> spasi biasa
        Return Regex.Replace(s, "\s+", " ").Trim()
    End Function

    ''' <summary>
    ''' Betulkan pola salah tulis yang sudah diketahui. Tiap perbaikan dicatat.
    ''' </summary>
    Public Function Repair(expr As String) As String
        Dim s = Normalize(expr)
        Dim before As String

        ' 1. kutip penutup lupa ditulis:  'teks, if(   ->  'teks', if(
        If (s.Length - s.Replace("'", "").Length) Mod 2 = 1 Then
            before = s
            s = Regex.Replace(s, "'([^']*?),\s*(if\s*\()", "'$1', $2", RegexOptions.IgnoreCase)
            If s <> before Then Say("    ~ kutip penutup ditambahkan: " & Clip(before) & "  ->  " & Clip(s))
        End If

        ' 2. koma antar argumen lupa ditulis:  'teks'if(  ->  'teks', if(
        before = s
        s = Regex.Replace(s, "('[^']*')\s*(if\s*\()", "$1, $2", RegexOptions.IgnoreCase)
        If s <> before Then Say("    ~ koma antar argumen ditambahkan: " & Clip(before) & "  ->  " & Clip(s))

        ' 3. .contains(...) sudah benar/salah -- '>0' tidak ada artinya
        before = s
        s = Regex.Replace(s, "(\.\s*(?:contains|includes)\s*\([^)]*\))\s*>\s*0", "$1", RegexOptions.IgnoreCase)
        If s <> before Then Say("    ~ '>0' setelah .contains() dibuang: " & Clip(before))

        ' 4. pemisah argumen ';' -> ','  (di luar kutip)
        If s.Contains(";") Then
            before = s
            s = ReplaceOutsideQuotes(s, ";"c, ","c)
            If s <> before Then Say("    ~ pemisah ';' diganti ',': " & Clip(before))
        End If

        ' 5. x = 'a' or 'b'   ->   x = 'a' or x = 'b'
        before = s
        s = Regex.Replace(s,
            "([A-Za-z_$][\w$.]*)\s*(==?)\s*('[^']*')\s+(or|\|\|)\s+('[^']*')(?!\s*[=<>!])",
            "$1 $2 $3 $4 $1 $2 $5", RegexOptions.IgnoreCase)
        If s <> before Then Say("    ~ pembandingan kedua dilengkapi: " & Clip(before) & "  ->  " & Clip(s))

        Return s
    End Function

    Private Shared Function ReplaceOutsideQuotes(s As String, oldCh As Char, newCh As Char) As String
        Dim sb As New StringBuilder(s.Length)
        Dim q As Char = ChrW(0)
        For Each ch In s
            If q = ChrW(0) AndAlso (ch = "'"c OrElse ch = """"c) Then
                q = ch
            ElseIf q <> ChrW(0) AndAlso ch = q Then
                q = ChrW(0)
            End If
            sb.Append(If(q = ChrW(0) AndAlso ch = oldCh, newCh, ch))
        Next
        Return sb.ToString()
    End Function

    Private Shared Function Clip(s As String) As String
        Return If(s.Length > 70, s.Substring(0, 70) & "…", s)
    End Function

    ' =====================================================  API

    ''' <summary>Evaluasi jadi nilai. Gagal -> Nothing (dicatat).</summary>
    Public Function EvalValue(expr As String) As Object
        Try
            _toks = Tokenize(Repair(expr))
            _pos = 0
            Dim v = ParseOr()
            If _pos < _toks.Count Then Throw New FormatException("sisa token '" & _toks(_pos) & "'")
            Return v
        Catch ex As Exception
            Say("    ! <<" & Clip(expr) & ">> : " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function EvalBool(expr As String) As Boolean
        Return Truthy(EvalValue(expr))
    End Function

    ''' <summary>Evaluasi lalu ubah jadi teks siap tampil.</summary>
    Public Function Render(expr As String) As String
        Return ToText(EvalValue(expr))
    End Function

    ''' <summary>True kalau nama paling depan ada di data (untuk deteksi penanda tak dikenal).</summary>
    Public Function KnowsRoot(expr As String) As Boolean
        Dim m = Regex.Match(Normalize(expr), "^([A-Za-z_$][\w$]*)")
        If Not m.Success Then Return False
        Return _data.ContainsKey(m.Groups(1).Value)
    End Function

    ' =====================================================  tokenizer

    Private Shared ReadOnly TokRx As New Regex(
        "\G(?:(?<s>'(?:[^'\\]|\\.)*'|""(?:[^""\\]|\\.)*"")" &
        "|(?<n>\d+(?:\.\d+)?)" &
        "|(?<i>[A-Za-z_$][\w$]*)" &
        "|(?<o>&&|\|\||==|!=|<>|<=|>=|[-+*/!<>=().,]))", RegexOptions.Compiled)

    Private Shared Function Tokenize(s As String) As List(Of String)
        Dim res As New List(Of String)
        Dim i = 0
        While i < s.Length
            If Char.IsWhiteSpace(s(i)) Then
                i += 1
                Continue While
            End If
            Dim m = TokRx.Match(s, i)
            If Not m.Success Then Throw New FormatException("karakter tak dikenal di '" & s.Substring(i) & "'")
            res.Add(m.Value)
            i = m.Index + m.Length
        End While
        Return res
    End Function

    ' =====================================================  parser

    Private Function Peek() As String
        Return If(_pos < _toks.Count, _toks(_pos), Nothing)
    End Function

    Private Function Take() As String
        Dim t = Peek()
        _pos += 1
        Return t
    End Function

    Private Function IsTok(t As String, ParamArray ops As String()) As Boolean
        If t Is Nothing Then Return False
        For Each o In ops
            If String.Equals(o, t, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    Private Function Accept(ParamArray ops As String()) As String
        If IsTok(Peek(), ops) Then Return Take()
        Return Nothing
    End Function

    Private Function ParseOr() As Object
        Dim v = ParseAnd()
        While Accept("||", "or") IsNot Nothing
            Dim r = ParseAnd()
            v = Truthy(v) OrElse Truthy(r)
        End While
        Return v
    End Function

    Private Function ParseAnd() As Object
        Dim v = ParseCmp()
        While Accept("&&", "and") IsNot Nothing
            Dim r = ParseCmp()
            v = Truthy(v) AndAlso Truthy(r)
        End While
        Return v
    End Function

    Private Function ParseCmp() As Object
        Dim l = ParseAdd()
        Dim op = Accept("==", "=", "!=", "<>", "<", "<=", ">", ">=")
        If op Is Nothing Then Return l
        Dim r = ParseAdd()
        Return Compare(l, r, op)
    End Function

    Private Function ParseAdd() As Object
        Dim v = ParseMul()
        While True
            Dim op = Accept("+", "-")
            If op Is Nothing Then Return v
            Dim r = ParseMul()
            If op = "+" AndAlso (TypeOf v Is String OrElse TypeOf r Is String) Then
                v = ToText(v) & ToText(r)
            Else
                v = If(op = "+", ToNum(v) + ToNum(r), ToNum(v) - ToNum(r))
            End If
        End While
        Return v
    End Function

    Private Function ParseMul() As Object
        Dim v = ParseUnary()
        While True
            Dim op = Accept("*", "/")
            If op Is Nothing Then Return v
            Dim r = ParseUnary()
            If op = "*" Then
                v = ToNum(v) * ToNum(r)
            Else
                Dim d = ToNum(r)
                v = If(d = 0, CObj(Nothing), CObj(ToNum(v) / d))
            End If
        End While
        Return v
    End Function

    Private Function ParseUnary() As Object
        If Accept("!", "not") IsNot Nothing Then Return Not Truthy(ParseUnary())
        If Accept("-") IsNot Nothing Then Return -ToNum(ParseUnary())
        Return ParsePostfix()
    End Function

    Private Function ParsePostfix() As Object
        Dim v = ParsePrimary()
        While IsTok(Peek(), ".")
            Take()
            Dim name = Take()
            If name Is Nothing Then Throw New FormatException("nama anggota kosong setelah '.'")
            If IsTok(Peek(), "(") Then
                Take()
                Dim args = ParseArgs()
                v = CallMethod(v, name, args)
            Else
                v = Member(v, name)
            End If
        End While
        Return v
    End Function

    Private Function ParseArgs() As List(Of Object)
        Dim args As New List(Of Object)
        If IsTok(Peek(), ")") Then
            Take()
            Return args
        End If
        Do
            args.Add(ParseOr())
        Loop While Accept(",") IsNot Nothing
        If Accept(")") Is Nothing Then Throw New FormatException("kurang ')'")
        Return args
    End Function

    Private Function ParsePrimary() As Object
        Dim t = Take()
        If t Is Nothing Then Throw New FormatException("ekspresi terpotong")

        If t.Length >= 2 AndAlso (t(0) = "'"c OrElse t(0) = """"c) Then
            Return t.Substring(1, t.Length - 2)
        End If

        Dim num As Double
        If Double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, num) Then Return num

        If t = "(" Then
            Dim v = ParseOr()
            If Accept(")") Is Nothing Then Throw New FormatException("kurang ')'")
            Return v
        End If

        Select Case t.ToLowerInvariant()
            Case "true" : Return True
            Case "false" : Return False
            Case "null" : Return Nothing
        End Select

        ' fungsi
        If IsTok(Peek(), "(") Then
            Take()
            Dim args = ParseArgs()
            Return CallFunction(t, args)
        End If

        ' variabel
        If _data.ContainsKey(t) Then Return _data(t)
        Dim hit = _data.Keys.FirstOrDefault(Function(k) String.Equals(k, t, StringComparison.OrdinalIgnoreCase))
        If hit IsNot Nothing Then Return _data(hit)
        Throw New FormatException("variabel '" & t & "' tidak ada di data")
    End Function

    ' =====================================================  fungsi & method

    Private Function CallFunction(name As String, args As List(Of Object)) As Object
        Select Case name.ToLowerInvariant()
            Case "if", "iif"
                If args.Count < 2 Then Throw New FormatException("if() butuh minimal 2 argumen")
                If Truthy(args(0)) Then Return args(1)
                Return If(args.Count > 2, args(2), CObj(""))
            Case "format"
                If args.Count < 2 Then Throw New FormatException("format() butuh 2 argumen")
                Return FormatValue(args(0), ToText(args(1)))
            Case "upper" : Return ToText(args(0)).ToUpperInvariant()
            Case "lower" : Return ToText(args(0)).ToLowerInvariant()
            Case Else
                Throw New FormatException("fungsi '" & name & "' tidak dikenal")
        End Select
    End Function

    Private Function CallMethod(target As Object, name As String, args As List(Of Object)) As Object
        Dim n = name.ToLowerInvariant()
        Select Case n
            Case "contains", "includes"
                Dim needle = If(args.Count > 0, ToText(args(0)), "")
                Dim lst = TryCast(target, List(Of Object))
                If lst IsNot Nothing Then
                    Return lst.Any(Function(x) String.Equals(ToText(x), needle, StringComparison.OrdinalIgnoreCase))
                End If
                Return ToText(target).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0
            Case "startswith" : Return ToText(target).StartsWith(ToText(args(0)), StringComparison.OrdinalIgnoreCase)
            Case "endswith" : Return ToText(target).EndsWith(ToText(args(0)), StringComparison.OrdinalIgnoreCase)
            Case "tolowercase" : Return ToText(target).ToLowerInvariant()
            Case "touppercase" : Return ToText(target).ToUpperInvariant()
            Case "trim" : Return ToText(target).Trim()
            Case Else : Throw New FormatException("method '" & name & "' tidak dikenal")
        End Select
    End Function

    Private Function Member(target As Object, name As String) As Object
        Dim n = name.ToLowerInvariant()
        Dim lst = TryCast(target, List(Of Object))
        If lst IsNot Nothing AndAlso (n = "length" OrElse n = "count") Then Return CDbl(lst.Count)
        If TypeOf target Is String AndAlso (n = "length" OrElse n = "count") Then
            Return CDbl(DirectCast(target, String).Length)
        End If
        Dim dic = TryCast(target, Dictionary(Of String, Object))
        If dic IsNot Nothing Then
            If dic.ContainsKey(name) Then Return dic(name)
            Dim hit = dic.Keys.FirstOrDefault(Function(k) String.Equals(k, name, StringComparison.OrdinalIgnoreCase))
            If hit IsNot Nothing Then Return dic(hit)
            Return Nothing
        End If
        Throw New FormatException("'" & name & "' bukan anggota yang dikenal")
    End Function

    ''' <summary>format(nilai, pola). Pola huruf d/M/y = tanggal, selain itu = angka.</summary>
    Private Function FormatValue(v As Object, pattern As String) As Object
        If v Is Nothing Then Return ""
        Dim p = pattern.Trim()
        Dim isDate = Regex.IsMatch(p, "[dMy]", RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(p, "[#0]")
        If isDate Then
            ' pola gaya Word: dd/mm/yyyy -> dd/MM/yyyy (mm di .NET = menit)
            Dim net = Regex.Replace(p, "m+", Function(mm) New String("M"c, mm.Value.Length))
            Dim dt As Date
            Dim txt = ToText(v)
            If Date.TryParse(txt, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) _
               OrElse Date.TryParseExact(txt, {"yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy"},
                                         CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                Return dt.ToString(net, CultureInfo.InvariantCulture)
            End If
            Say("    ! format(): '" & txt & "' bukan tanggal yang bisa dibaca")
            Return txt
        End If
        Return ToNum(v).ToString(p, CultureInfo.InvariantCulture)
    End Function

    ' =====================================================  nilai

    Public Shared Function Truthy(v As Object) As Boolean
        If v Is Nothing Then Return False
        If TypeOf v Is Boolean Then Return DirectCast(v, Boolean)
        If TypeOf v Is Double Then Return DirectCast(v, Double) <> 0
        If TypeOf v Is String Then Return DirectCast(v, String).Length > 0
        Dim lst = TryCast(v, List(Of Object))
        If lst IsNot Nothing Then Return lst.Count > 0
        Dim dic = TryCast(v, Dictionary(Of String, Object))
        If dic IsNot Nothing Then Return dic.Count > 0
        Return True
    End Function

    Public Shared Function ToText(v As Object) As String
        If v Is Nothing Then Return ""
        If TypeOf v Is Boolean Then Return If(DirectCast(v, Boolean), "true", "false")
        If TypeOf v Is Double Then
            Dim d = DirectCast(v, Double)
            If d = Math.Floor(d) AndAlso Math.Abs(d) < 1.0E+15 Then
                Return CLng(d).ToString(CultureInfo.InvariantCulture)
            End If
            Return d.ToString(CultureInfo.InvariantCulture)
        End If
        If TypeOf v Is String Then Return DirectCast(v, String)
        Dim lst = TryCast(v, List(Of Object))
        If lst IsNot Nothing Then Return String.Join(", ", lst.Select(AddressOf ToText))
        Return v.ToString()
    End Function

    Private Shared Function ToNum(v As Object) As Double
        If v Is Nothing Then Return 0
        If TypeOf v Is Double Then Return DirectCast(v, Double)
        If TypeOf v Is Boolean Then Return If(DirectCast(v, Boolean), 1, 0)
        Dim d As Double
        If Double.TryParse(ToText(v), NumberStyles.Any, CultureInfo.InvariantCulture, d) Then Return d
        Return 0
    End Function

    Private Shared Function Compare(l As Object, r As Object, op As String) As Object
        Dim ln, rn As Double
        Dim bothNum = Double.TryParse(ToText(l), NumberStyles.Any, CultureInfo.InvariantCulture, ln) AndAlso
                      Double.TryParse(ToText(r), NumberStyles.Any, CultureInfo.InvariantCulture, rn)
        Dim c As Integer
        If bothNum Then
            c = ln.CompareTo(rn)
        Else
            c = String.Compare(ToText(l), ToText(r), StringComparison.OrdinalIgnoreCase)
        End If
        Select Case op
            Case "==", "=" : Return c = 0
            Case "!=", "<>" : Return c <> 0
            Case "<" : Return c < 0
            Case "<=" : Return c <= 0
            Case ">" : Return c > 0
            Case ">=" : Return c >= 0
        End Select
        Return False
    End Function

End Class
