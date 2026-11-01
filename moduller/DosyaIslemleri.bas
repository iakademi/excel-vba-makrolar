Attribute VB_Name = "DosyaIslemleri"
'==============================================================
' Modul   : DosyaIslemleri
' Amac    : Klasor tarama, calisma kitabi birlestirme, PDF disa aktarma
'==============================================================
Option Explicit

' Kullaniciya klasor sectirir; iptal edilirse bos metin doner
Private Function KlasorSec(ByVal baslik As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = baslik
        If .Show = -1 Then
            KlasorSec = .SelectedItems(1)
            If Right$(KlasorSec, 1) <> "\" Then KlasorSec = KlasorSec & "\"
        End If
    End With
End Function

' Secilen klasordeki dosyalari yeni bir sayfaya listeler (ad, boyut, degistirilme tarihi)
Public Sub KlasorDosyalariniListele()
    Dim klasor As String
    Dim dosya As String
    Dim ws As Worksheet
    Dim satir As Long

    klasor = KlasorSec("Listelenecek klasoru secin")
    If klasor = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Range("A1:D1").Value = Array("Dosya Adi", "Boyut (KB)", "Degistirilme", "Tam Yol")
    ws.Range("A1:D1").Font.Bold = True

    satir = 2
    dosya = Dir(klasor & "*.*")
    Do While dosya <> ""
        ws.Cells(satir, 1).Value = dosya
        ws.Cells(satir, 2).Value = Round(FileLen(klasor & dosya) / 1024, 1)
        ws.Cells(satir, 3).Value = FileDateTime(klasor & dosya)
        ws.Cells(satir, 4).Value = klasor & dosya
        satir = satir + 1
        dosya = Dir()
    Loop

    ws.Columns("A:D").AutoFit
    ws.Columns("C").NumberFormat = "dd.mm.yyyy hh:mm"
    MsgBox (satir - 2) & " dosya listelendi.", vbInformation
End Sub

' Secilen klasordeki tum .xlsx dosyalarinin ilk sayfasini bu kitaba alt alta birlestirir.
' Her dosyanin basliklari ayni sirada ve ilk satirda olmalidir.
Public Sub CalismaKitaplariniBirlestir()
    Dim klasor As String
    Dim dosya As String
    Dim kaynakKitap As Workbook
    Dim hedef As Worksheet
    Dim kaynakSayfa As Worksheet
    Dim sonSatir As Long
    Dim hedefSatir As Long
    Dim ilk As Boolean
    Dim adet As Long

    klasor = KlasorSec("Birlestirilecek .xlsx dosyalarinin klasoru")
    If klasor = "" Then Exit Sub

    On Error GoTo Hata
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set hedef = ThisWorkbook.Worksheets.Add
    hedef.Name = "Birlesik_" & Format(Now, "hhnnss")
    ilk = True
    hedefSatir = 1

    dosya = Dir(klasor & "*.xlsx")
    Do While dosya <> ""
        ' Bu kitabin kendisini atla
        If dosya <> ThisWorkbook.Name Then
            Set kaynakKitap = Workbooks.Open(klasor & dosya, ReadOnly:=True)
            Set kaynakSayfa = kaynakKitap.Worksheets(1)
            sonSatir = kaynakSayfa.Cells(kaynakSayfa.Rows.Count, 1).End(xlUp).Row

            If sonSatir >= 1 And Application.WorksheetFunction.CountA(kaynakSayfa.UsedRange) > 0 Then
                If ilk Then
                    ' Ilk dosyada basliklar dahil kopyala; son sutuna kaynak dosya adi yazilir
                    kaynakSayfa.UsedRange.Copy hedef.Cells(1, 1)
                    hedef.Cells(1, kaynakSayfa.UsedRange.Columns.Count + 1).Value = "Kaynak Dosya"
                    hedefSatir = sonSatir + 1
                    hedef.Range(hedef.Cells(2, kaynakSayfa.UsedRange.Columns.Count + 1), _
                                hedef.Cells(sonSatir, kaynakSayfa.UsedRange.Columns.Count + 1)).Value = dosya
                    ilk = False
                Else
                    ' Sonraki dosyalarda baslik satirini atla
                    If sonSatir >= 2 Then
                        kaynakSayfa.Range(kaynakSayfa.Rows(2), kaynakSayfa.Rows(sonSatir)).Copy _
                            hedef.Cells(hedefSatir, 1)
                        hedef.Range(hedef.Cells(hedefSatir, kaynakSayfa.UsedRange.Columns.Count + 1), _
                                    hedef.Cells(hedefSatir + sonSatir - 2, kaynakSayfa.UsedRange.Columns.Count + 1)).Value = dosya
                        hedefSatir = hedefSatir + sonSatir - 1
                    End If
                End If
                adet = adet + 1
            End If
            kaynakKitap.Close SaveChanges:=False
        End If
        dosya = Dir()
    Loop

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    hedef.Columns.AutoFit
    MsgBox adet & " dosya birlestirildi.", vbInformation
    Exit Sub
Hata:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Kitaptaki gorunur sayfalarin her birini ayri PDF dosyasi olarak kaydeder
Public Sub SayfalariPdfOlarakKaydet()
    Dim klasor As String
    Dim ws As Worksheet
    Dim yol As String
    Dim adet As Long

    klasor = KlasorSec("PDF dosyalarinin kaydedilecegi klasoru secin")
    If klasor = "" Then Exit Sub

    On Error GoTo Hata
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            yol = klasor & GecerliDosyaAdi(ws.Name) & ".pdf"
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, _
                Quality:=xlQualityStandard, OpenAfterPublish:=False
            adet = adet + 1
        End If
    Next ws
    MsgBox adet & " PDF olusturuldu: " & klasor, vbInformation
    Exit Sub
Hata:
    MsgBox "PDF olusturulamadi: " & Err.Description, vbCritical
End Sub

' Dosya adinda kullanilamayan karakterleri alt cizgi ile degistirir
Private Function GecerliDosyaAdi(ByVal ad As String) As String
    Dim yasakli As Variant
    Dim k As Variant
    yasakli = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each k In yasakli
        ad = Replace(ad, k, "_")
    Next k
    GecerliDosyaAdi = ad
End Function
