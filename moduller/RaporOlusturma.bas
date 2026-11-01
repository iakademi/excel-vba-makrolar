Attribute VB_Name = "RaporOlusturma"
'==============================================================
' Modul   : RaporOlusturma
' Amac    : Veri sayfasindan ozet rapor sayfasi ve grafik uretmek
' Beklenen veri duzeni (aktif sayfada, A1'den baslayan tablo):
'   1. satir basliklar; A sutunu = kategori (metin), B sutunu = tutar (sayi)
'==============================================================
Option Explicit

' Aktif sayfadaki veriyi kategoriye gore toplayip "Rapor" sayfasina yazar
' ve bir sutun grafigi ekler
Public Sub OzetRaporOlustur()
    Dim kaynak As Worksheet
    Dim rapor As Worksheet
    Dim sozluk As Object
    Dim sonSatir As Long
    Dim i As Long
    Dim anahtar As Variant
    Dim satir As Long
    Dim grafik As ChartObject

    Set kaynak = ActiveSheet
    sonSatir = kaynak.Cells(kaynak.Rows.Count, "A").End(xlUp).Row
    If sonSatir < 2 Then
        MsgBox "Ozetlenecek veri bulunamadi.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Hata
    Application.ScreenUpdating = False

    ' Kategorilere gore toplam: Scripting.Dictionary (Windows)
    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = vbTextCompare
    For i = 2 To sonSatir
        If IsNumeric(kaynak.Cells(i, "B").Value) Then
            anahtar = Trim$(CStr(kaynak.Cells(i, "A").Value))
            sozluk(anahtar) = sozluk(anahtar) + CDbl(kaynak.Cells(i, "B").Value)
        End If
    Next i

    ' Eski rapor sayfasini sil, yenisini olustur
    Set rapor = SayfaOlustur("Rapor")

    rapor.Range("A1").Value = "Kategori Ozet Raporu"
    rapor.Range("A1").Font.Bold = True
    rapor.Range("A1").Font.Size = 14
    rapor.Range("A2").Value = "Olusturma zamani: " & Format(Now, "dd.mm.yyyy hh:nn")

    rapor.Range("A4:C4").Value = Array("Kategori", "Toplam", "Pay (%)")
    With rapor.Range("A4:C4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With

    satir = 5
    For Each anahtar In sozluk.Keys
        rapor.Cells(satir, 1).Value = anahtar
        rapor.Cells(satir, 2).Value = sozluk(anahtar)
        satir = satir + 1
    Next anahtar

    If satir = 5 Then
        MsgBox "Sayisal veri bulunamadi.", vbExclamation
        GoTo Cikis
    End If

    ' Pay sutunu icin formuller ve toplam satiri
    rapor.Range("C5:C" & satir - 1).Formula = "=B5/SUM($B$5:$B$" & satir - 1 & ")"
    rapor.Cells(satir, 1).Value = "Genel Toplam"
    rapor.Cells(satir, 2).Formula = "=SUM(B5:B" & satir - 1 & ")"
    rapor.Rows(satir).Font.Bold = True

    rapor.Range("B5:B" & satir).NumberFormat = "#,##0.00"
    rapor.Range("C5:C" & satir - 1).NumberFormat = "0.0%"
    rapor.Columns("A:C").AutoFit

    ' Grafik
    Set grafik = rapor.ChartObjects.Add(Left:=rapor.Range("E4").Left, _
                                        Top:=rapor.Range("E4").Top, _
                                        Width:=420, Height:=250)
    With grafik.Chart
        .ChartType = xlColumnClustered
        .SetSourceData rapor.Range("A4:B" & satir - 1)
        .HasTitle = True
        .ChartTitle.Text = "Kategoriye Gore Toplam"
        .HasLegend = False
    End With

    rapor.Activate
    rapor.Range("A1").Select
Cikis:
    Application.ScreenUpdating = True
    Exit Sub
Hata:
    Application.ScreenUpdating = True
    MsgBox "Rapor olusturulamadi: " & Err.Description, vbCritical
End Sub

' Calisma kitabindaki tum sayfalari baglantili bir icerik tablosu olarak listeler
Public Sub SayfaIcerikTablosu()
    Dim ws As Worksheet
    Dim icerik As Worksheet
    Dim satir As Long

    On Error GoTo Hata
    Set icerik = SayfaOlustur("Icerik")
    icerik.Range("A1").Value = "Sayfa Icerigi"
    icerik.Range("A1").Font.Bold = True

    satir = 3
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> icerik.Name Then
            ' Her sayfa adina tiklanabilir baglanti ekle
            icerik.Hyperlinks.Add Anchor:=icerik.Cells(satir, 1), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            satir = satir + 1
        End If
    Next ws
    icerik.Columns("A").AutoFit
    Exit Sub
Hata:
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Verilen adda bos bir sayfa olusturur; varsa silip yeniden olusturur
Private Function SayfaOlustur(ByVal ad As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(ad).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set SayfaOlustur = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    SayfaOlustur.Name = ad
End Function
