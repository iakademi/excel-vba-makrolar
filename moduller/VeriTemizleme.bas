Attribute VB_Name = "VeriTemizleme"
'==============================================================
' Modul   : VeriTemizleme
' Amac    : Secili aralikta yaygin veri temizleme islemleri
' Kullanim: Temizlenecek araligi secin, Alt+F8 ile makroyu calistirin
'==============================================================
Option Explicit

' Hizli calisma icin ekran guncellemesini ve hesaplamayi kapatir/acar
Private Sub HizModu(ByVal ac As Boolean)
    With Application
        .ScreenUpdating = Not ac
        .EnableEvents = Not ac
        .Calculation = IIf(ac, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub

' Secili hucrelerdeki bas/son/fazla bosluklari ve gorunmez karakterleri temizler
Public Sub BosluklariTemizle()
    Dim hucre As Range
    Dim aralik As Range
    Dim sayac As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Lutfen once bir hucre araligi secin.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Hata
    HizModu True

    ' Yalnizca sabit metin iceren hucreleri isle (formullere dokunma)
    On Error Resume Next
    Set aralik = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo Hata

    If Not aralik Is Nothing Then
        For Each hucre In aralik.Cells
            Dim eski As String, yeni As String
            eski = CStr(hucre.Value)
            yeni = Replace(eski, Chr(160), " ")       ' bolunmez bosluk
            yeni = Replace(yeni, vbTab, " ")           ' sekme
            yeni = Replace(yeni, vbCr, "")             ' satir sonlari
            yeni = Replace(yeni, vbLf, " ")
            yeni = Application.WorksheetFunction.Trim(yeni) ' fazla bosluklari teke indir
            If yeni <> eski Then
                hucre.Value = yeni
                sayac = sayac + 1
            End If
        Next hucre
    End If

    HizModu False
    MsgBox sayac & " hucre temizlendi.", vbInformation
    Exit Sub
Hata:
    HizModu False
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Secili aralikta tamamen bos olan satirlari siler
Public Sub BosSatirlariSil()
    Dim aralik As Range
    Dim satir As Long
    Dim silinen As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set aralik = Selection

    If MsgBox("Secili aralikta tamamen bos satirlar silinecek. Devam edilsin mi?", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    On Error GoTo Hata
    HizModu True

    ' Silme isleminde kayma olmamasi icin asagidan yukariya dogru git
    For satir = aralik.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(aralik.Rows(satir)) = 0 Then
            aralik.Rows(satir).EntireRow.Delete
            silinen = silinen + 1
        End If
    Next satir

    HizModu False
    MsgBox silinen & " bos satir silindi.", vbInformation
    Exit Sub
Hata:
    HizModu False
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Secili aralikta secilen sutunlara gore tekrar eden satirlari kaldirir
' (ilk sutun anahtar kabul edilir; birden fazla sutun icin Columns dizisini degistirin)
Public Sub TekrarEdenleriSil()
    If TypeName(Selection) <> "Range" Then Exit Sub

    On Error GoTo Hata
    ' Header:=xlYes -> ilk satir baslik satiri
    Selection.RemoveDuplicates Columns:=1, Header:=xlYes
    MsgBox "Tekrar eden kayitlar kaldirildi (ilk sutuna gore).", vbInformation
    Exit Sub
Hata:
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Metin bicimini duzenler: 1 = BUYUK HARF, 2 = kucuk harf, 3 = Her Kelimenin Ilk Harfi Buyuk
' Turkce i/I ve ı/İ donusumlerini dogru yapmak icin elle esleme kullanilir
Public Sub BuyukKucukHarfDuzenle()
    Dim secim As String
    Dim hucre As Range
    Dim aralik As Range

    secim = InputBox("1 = BUYUK HARF" & vbCrLf & "2 = kucuk harf" & vbCrLf & _
                     "3 = Her Kelimenin Ilk Harfi Buyuk", "Harf Duzeni", "3")
    If secim <> "1" And secim <> "2" And secim <> "3" Then Exit Sub

    On Error Resume Next
    Set aralik = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If aralik Is Nothing Then Exit Sub

    For Each hucre In aralik.Cells
        Select Case secim
            Case "1": hucre.Value = TrBuyuk(CStr(hucre.Value))
            Case "2": hucre.Value = TrKucuk(CStr(hucre.Value))
            Case "3": hucre.Value = TrBaslik(CStr(hucre.Value))
        End Select
    Next hucre
End Sub

' --- Turkce duyarli donusum yardimcilari ---
Private Function TrBuyuk(ByVal s As String) As String
    s = Replace(s, "i", ChrW(304))        ' i -> İ
    s = Replace(s, ChrW(305), "I")        ' ı -> I
    TrBuyuk = UCase$(s)
End Function

Private Function TrKucuk(ByVal s As String) As String
    s = Replace(s, "I", ChrW(305))        ' I -> ı
    s = Replace(s, ChrW(304), "i")        ' İ -> i
    TrKucuk = LCase$(s)
End Function

Private Function TrBaslik(ByVal s As String) As String
    Dim kelimeler() As String
    Dim i As Long
    kelimeler = Split(TrKucuk(s), " ")
    For i = LBound(kelimeler) To UBound(kelimeler)
        If Len(kelimeler(i)) > 0 Then
            kelimeler(i) = TrBuyuk(Left$(kelimeler(i), 1)) & Mid$(kelimeler(i), 2)
        End If
    Next i
    TrBaslik = Join(kelimeler, " ")
End Function
