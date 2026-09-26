Option Strict On
Option Infer On

Imports System.Text

' ============================================================================
'  GlobalSettings.vb  -- kerangka page.css (satu file untuk semua HTML).
'  Ukuran halaman & margin diambil dari w:sectPr dokumen, bukan angka tetap.
' ============================================================================

Public Module GlobalSettings

    Public Const CssFileName As String = "page.css"

    ''' <summary>
    ''' Bagian tetap page.css. Class hasil dedup (p*, r*, t*, c*, tr*)
    ''' ditempel setelahnya oleh Generator.
    ''' </summary>
    Public Function BuildPageCss(m As DocModel) As String
        Dim s = m.Setup
        Dim pageW = Cm(s.WidthTw)
        Dim pageH = Cm(s.HeightTw)
        Dim mT = Cm(s.MarginTopTw)
        Dim mB = Cm(s.MarginBottomTw)
        Dim mL = Cm(s.MarginLeftTw)
        Dim mR = Cm(s.MarginRightTw)
        Dim hD = Cm(s.HeaderTw)
        Dim fD = Cm(s.FooterTw)
        Dim contentW = Cm(s.ContentWidthTw)

        Dim sb As New StringBuilder()
        sb.AppendLine("/* page.css -- dihasilkan RiplayWord2Html dari .docx")
        sb.AppendLine("   Dipakai bersama oleh header.html, footer.html dan semua bodyN.html. */")
        sb.AppendLine()

        sb.AppendLine("@page { size: " & pageW & " " & pageH & "; margin: 0; }")
        sb.AppendLine()

        sb.AppendLine("* { box-sizing: border-box; }")
        sb.AppendLine("html, body { margin: 0; padding: 0; background: #fff; }")
        sb.AppendLine("body {")
        sb.AppendLine("  font-family: '" & m.DefaultFont & "', Arial, sans-serif;")
        sb.AppendLine("  font-size: " & Pt(m.DefaultSizeHalfPt) & ";")
        sb.AppendLine("  color: #000;")
        sb.AppendLine("  line-height: 1.16;")
        sb.AppendLine("  -webkit-print-color-adjust: exact;")
        sb.AppendLine("  print-color-adjust: exact;")
        sb.AppendLine("}")
        sb.AppendLine()

        sb.AppendLine("p { margin: 0; }")
        sb.AppendLine("sup, sub { font-size: 65%; line-height: 0; }")
        sb.AppendLine("img { display: inline-block; vertical-align: top; }")
        sb.AppendLine()

        sb.AppendLine("/* tabel mengikuti lebar grid Word */")
        sb.AppendLine("table { border-collapse: collapse; table-layout: fixed; border-spacing: 0; }")
        sb.AppendLine("td { padding: 0; vertical-align: top; overflow-wrap: break-word; }")
        sb.AppendLine()

        sb.AppendLine("/* label nomor/huruf pada paragraf hanging-indent (meniru tab stop Word) */")
        sb.AppendLine(".lbl { display: inline-block; text-indent: 0; }")
        sb.AppendLine(".tab { display: inline-block; width: 0.5cm; }")
        sb.AppendLine()

        sb.AppendLine("/* lebar area isi = lebar halaman dikurangi margin kiri-kanan */")
        sb.AppendLine(".hdr, .ftr, .bdy { width: " & contentW & "; }")
        sb.AppendLine()

        sb.AppendLine("/* tampilan saat file dibuka sendiri-sendiri */")
        sb.AppendLine("body > .hdr, body > .ftr, body > .bdy { margin: " & mT & " auto; }")
        sb.AppendLine()

        sb.AppendLine("/* kerangka halaman utuh (dipakai AllPages.html) */")
        sb.AppendLine(".page {")
        sb.AppendLine("  position: relative;")
        sb.AppendLine("  width: " & pageW & ";")
        sb.AppendLine("  height: " & pageH & ";")
        sb.AppendLine("  padding: " & mT & " " & mR & " " & mB & " " & mL & ";")
        sb.AppendLine("  overflow: hidden;")
        sb.AppendLine("  background: #fff;")
        sb.AppendLine("}")
        sb.AppendLine(".page + .page { page-break-before: always; break-before: page; }")
        sb.AppendLine(".page > .hdr { position: absolute; top: " & hD & "; left: " & mL & "; right: " & mR & "; }")
        sb.AppendLine(".page > .ftr { position: absolute; bottom: " & fD & "; left: " & mL & "; right: " & mR & "; }")
        sb.AppendLine(".page > .bdy { width: auto; }")
        sb.AppendLine()

        sb.AppendLine(".page-break { break-after: page; page-break-after: always; height: 0; }")
        sb.AppendLine()

        sb.AppendLine("@media screen {")
        sb.AppendLine("  body { background: #e9e9ec; }")
        sb.AppendLine("  .page { margin: 0.6cm auto; box-shadow: 0 0 6px rgba(0,0,0,.28); }")
        sb.AppendLine("  body > .hdr, body > .ftr, body > .bdy { background: #fff; padding: 0.5cm; }")
        sb.AppendLine("}")
        sb.AppendLine()

        sb.AppendLine("/* ---- class hasil konversi format Word ---- */")
        Return sb.ToString()
    End Function

End Module
