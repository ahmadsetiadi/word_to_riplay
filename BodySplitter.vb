Option Strict On
Option Infer On

Imports System.Text.RegularExpressions

' ============================================================================
'  BodySplitter.vb  -- pecah body jadi beberapa bagian.
'  Aturan: setiap paragraf yang isinya penanda "<<Page Break>>" menutup bagian
'  berjalan dan membuka bagian baru. Paragraf penanda itu sendiri dibuang.
' ============================================================================

Public Module BodySplitter

    ''' <summary>Cocok untuk "&lt;&lt;Page Break&gt;&gt;" tanpa peduli spasi/kapital.</summary>
    Private ReadOnly Marker As New Regex("^\s*<<\s*page\s*break\s*>>\s*$",
                                         RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    Public Function IsMarker(b As DocBlock) As Boolean
        If Not b.IsPara Then Return False
        Return Marker.IsMatch(b.Para.PlainText)
    End Function

    ''' <summary>
    ''' Pecah daftar blok pada setiap penanda. Bagian kosong di ujung dibuang,
    ''' supaya penanda terakhir tidak menghasilkan body kosong.
    ''' </summary>
    Public Function Split(blocks As List(Of DocBlock)) As List(Of List(Of DocBlock))
        Dim parts As New List(Of List(Of DocBlock))
        Dim cur As New List(Of DocBlock)

        For Each b In blocks
            If IsMarker(b) Then
                parts.Add(cur)
                cur = New List(Of DocBlock)
            Else
                cur.Add(b)
            End If
        Next
        parts.Add(cur)

        ' buang bagian yang tidak punya isi sama sekali
        Dim res As New List(Of List(Of DocBlock))
        For Each p In parts
            If HasContent(p) Then res.Add(p)
        Next
        Return res
    End Function

    Private Function HasContent(blocks As List(Of DocBlock)) As Boolean
        For Each b In blocks
            If b.IsTable Then Return True
            If b.Para.IsPageBreakOnly Then Continue For
            If b.Para.PlainText.Trim().Length > 0 Then Return True
            For Each r In b.Para.Runs
                If r.Kind = RunKind.Picture Then Return True
            Next
        Next
        Return False
    End Function

    ''' <summary>
    ''' Pecah satu body jadi beberapa halaman Word (dipotong di hard page break).
    ''' Dipakai AllPages.html supaya jumlah halaman preview sama dengan Word.
    ''' </summary>
    Public Function SplitAtPageBreaks(blocks As List(Of DocBlock)) As List(Of List(Of DocBlock))
        Dim pages As New List(Of List(Of DocBlock))
        Dim cur As New List(Of DocBlock)
        For Each b In blocks
            If b.IsPara AndAlso b.Para.IsPageBreakOnly Then
                pages.Add(cur)
                cur = New List(Of DocBlock)
            Else
                cur.Add(b)
            End If
        Next
        pages.Add(cur)

        Dim res As New List(Of List(Of DocBlock))
        For Each p In pages
            If p.Count > 0 Then res.Add(p)
        Next
        If res.Count = 0 Then res.Add(New List(Of DocBlock))
        Return res
    End Function

    ''' <summary>Buang page-break Word yang nyangkut di awal/akhir sebuah bagian.</summary>
    Public Sub TrimEdgeBreaks(blocks As List(Of DocBlock))
        While blocks.Count > 0 AndAlso blocks(0).IsPara AndAlso blocks(0).Para.IsPageBreakOnly
            blocks.RemoveAt(0)
        End While
        While blocks.Count > 0 AndAlso blocks(blocks.Count - 1).IsPara _
              AndAlso blocks(blocks.Count - 1).Para.IsPageBreakOnly
            blocks.RemoveAt(blocks.Count - 1)
        End While
    End Sub

End Module
