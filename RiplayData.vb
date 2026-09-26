Option Strict On
Option Infer On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Encodings.Web
Imports System.Text.Json

' ============================================================================
'  RiplayData.vb -- SEMUA variabel yang dikirim VB ke HTML.
'
'  Build()  : isi semua variabel (di sinilah Anda menyambungkan data aplikasi).
'             Key dictionary = nama yang ditulis di Word, mis. <<insuredname>>.
'  Write()  : tulis ke Output\<nama>\data.json
'  Read()   : baca kembali data.json (kalau mau generate ulang tanpa import lagi)
'
'  Mengganti isi Build() TIDAK memerlukan perubahan di Word maupun di HTML.
' ============================================================================

Public Class RiplayData

    Public Const FileName As String = "data.json"

    Private Shared ReadOnly Utf8NoBom As New UTF8Encoding(False)
    Private Shared ReadOnly JsonOpts As New JsonSerializerOptions With {
        .WriteIndented = True,
        .Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping}

    ' ------------------------------------------------------------------
    '  Kode yang dipakai di penanda Word
    '    currency : '88' = IDR   '02' = USD   '03' = SGD
    '    plancode : 'BMUC3' = 3 tahun, 'BMUC5' = 5 tahun, selain itu = sekaligus
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' CONTOH DATA. Ganti isi fungsi ini dengan data asli dari aplikasi.
    ''' Semua nilai uang sudah diformat di sini supaya Word cukup menulis
    ''' &lt;&lt;amountpremium&gt;&gt; tanpa perlu tahu aturan pemisah desimal.
    ''' </summary>
    Public Shared Function Build() As Dictionary(Of String, Object)
        Dim currency = "88"                    ' 88 = Rupiah
        Dim plancode = "BMUC5"                 ' 5 tahun
        Dim insuredEntryAge = 35
        Dim basicPremium As Double = 120000000
        Dim riderPremium As Double = 4500000
        Dim riders = New List(Of Object) From {"MiPayor"}

        Dim d As New Dictionary(Of String, Object)

        ' --- identitas -------------------------------------------------
        d("policyholdername") = "Siti Rahayu"
        d("policyholdersex") = "Perempuan"
        d("policyholderdob") = "1985-11-03"
        d("policyholderentryage") = 40

        d("insuredname") = "Budi Santoso"
        d("insuredsex") = "Laki-laki"
        d("insureddob") = "1990-05-17"
        d("insuredentryage") = insuredEntryAge

        ' --- polis -----------------------------------------------------
        d("currency") = currency
        d("plancode") = plancode
        d("policyterm") = 120 - insuredEntryAge

        ' --- premi & manfaat (sudah diformat sesuai mata uang) ---------
        d("amountpremium") = Money(basicPremium, currency)
        d("riderpremium") = Money(riderPremium, currency)
        d("totalpremium") = Money(basicPremium + riderPremium, currency)
        d("danamapan") = Money(basicPremium * 5 * 1.03, currency)

        ' --- rider -----------------------------------------------------
        d("riders") = riders
        d("ridercount") = riders.Count

        ' --- footer ----------------------------------------------------
        d("isname") = "Andi Wijaya"                  ' <<isname>>  tenaga pemasar
        d("iscode") = "AG-00871"                     ' <<iscode>>
        d("illustrationnumber") = "RP-2026-000123"   ' <<illustrationnumber>>
        d("illustrationdate") = Date.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        d("illustrationexpireddate") = Date.Today.AddDays(30).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)

        ' --- tabel ilustrasi (untuk <<illus.xxx>> di halaman 7) --------
        d("illus") = SampleIllustration(insuredEntryAge, basicPremium, currency)

        Return d
    End Function

    ''' <summary>Format uang mengikuti mata uang polis.</summary>
    Public Shared Function Money(v As Double, currency As String) As String
        Select Case currency
            Case "02" : Return "US$" & v.ToString("#,##0.00", CultureInfo.InvariantCulture)
            Case "03" : Return "S$" & v.ToString("#,##0.00", CultureInfo.InvariantCulture)
            Case Else
                ' Rupiah: ribuan '.', desimal ','
                Dim id = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                id.NumberFormat.NumberGroupSeparator = "."
                id.NumberFormat.NumberDecimalSeparator = ","
                Return "Rp" & v.ToString("#,##0.00", id)
        End Select
    End Function

    Private Shared Function SampleIllustration(entryAge As Integer, premi As Double,
                                               currency As String) As List(Of Object)
        Dim res As New List(Of Object)
        For i = 1 To 5
            Dim row As New Dictionary(Of String, Object) From {
                {"age", entryAge + i},
                {"year", i},
                {"premium", Money(premi, currency)},
                {"cashPayment", Money(premi * 0.05 * i, currency)},
                {"addCashPayment", Money(premi * 0.01 * i, currency)},
                {"totalCashPayment", Money(premi * 0.06 * i, currency)},
                {"cashValue", Money(premi * 0.9 * i, currency)},
                {"addCashValue", Money(premi * 0.05 * i, currency)},
                {"totalCashValue", Money(premi * 0.95 * i, currency)},
                {"adb", Money(premi, currency)}
            }
            res.Add(row)
        Next
        Return res
    End Function

    ' ------------------------------------------------------------------ IO

    Public Shared Function ToJson(data As IDictionary(Of String, Object)) As String
        Return JsonSerializer.Serialize(data, JsonOpts)
    End Function

    ''' <summary>Tulis data.json ke folder output. Mengembalikan path file.</summary>
    Public Shared Function Write(outDir As String, data As IDictionary(Of String, Object)) As String
        Dim p = Path.Combine(outDir, FileName)
        File.WriteAllText(p, ToJson(data), Utf8NoBom)
        Return p
    End Function

    ''' <summary>Baca data.json; kalau tidak ada, pakai Build().</summary>
    Public Shared Function Read(outDir As String) As Dictionary(Of String, Object)
        Dim p = Path.Combine(outDir, FileName)
        If Not File.Exists(p) Then Return Build()
        Using doc = JsonDocument.Parse(File.ReadAllText(p))
            Return CType(FromJson(doc.RootElement), Dictionary(Of String, Object))
        End Using
    End Function

    ''' <summary>JsonElement -> Dictionary / List / String / Double / Boolean / Nothing.</summary>
    Public Shared Function FromJson(el As JsonElement) As Object
        Select Case el.ValueKind
            Case JsonValueKind.Object
                Dim d As New Dictionary(Of String, Object)
                For Each pr In el.EnumerateObject()
                    d(pr.Name) = FromJson(pr.Value)
                Next
                Return d
            Case JsonValueKind.Array
                Dim l As New List(Of Object)
                For Each it In el.EnumerateArray()
                    l.Add(FromJson(it))
                Next
                Return l
            Case JsonValueKind.String
                Return el.GetString()
            Case JsonValueKind.Number
                Return el.GetDouble()
            Case JsonValueKind.True
                Return True
            Case JsonValueKind.False
                Return False
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>Normalkan hasil Build() lewat JSON supaya tipenya seragam
    ''' (angka jadi Double, objek jadi Dictionary) — sama seperti kalau dibaca dari file.</summary>
    Public Shared Function Normalize(data As IDictionary(Of String, Object)) As Dictionary(Of String, Object)
        Using doc = JsonDocument.Parse(ToJson(data))
            Return CType(FromJson(doc.RootElement), Dictionary(Of String, Object))
        End Using
    End Function

End Class
