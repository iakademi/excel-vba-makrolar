Attribute VB_Name = "MailGonderme"
'==============================================================
' Modul   : MailGonderme
' Amac    : Outlook uzerinden e-posta hazirlamak veya gondermek
' Gereksinim: Masaustu Microsoft Outlook kurulu ve yapilandirilmis olmali
' Gec baglama (late binding) kullanildigi icin ek referans gerekmez.
'
' Beklenen veri duzeni (aktif sayfa, 1. satir baslik):
'   A = Ad Soyad, B = E-posta, C = Konu, D = Mesaj metni
'==============================================================
Option Explicit

' Outlook uygulamasina baglanir; acik degilse baslatir
Private Function OutlookAl() As Object
    Dim ol As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    Set OutlookAl = ol
End Function

' Basit e-posta bicim kontrolu (tam dogrulama degildir)
Private Function EpostaGecerliMi(ByVal adres As String) As Boolean
    adres = Trim$(adres)
    EpostaGecerliMi = (adres Like "?*@?*.?*") And InStr(adres, " ") = 0
End Function

' Secili satirlar icin kisisellestirilmis e-posta TASLAKLARI acar (otomatik gondermez).
' Gonderim oncesi her taslagi gozden gecirebilirsiniz.
Public Sub SeciliSatirlaraMailHazirla()
    Dim ol As Object
    Dim posta As Object
    Dim satir As Range
    Dim ws As Worksheet
    Dim adet As Long
    Dim atlanan As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Lutfen gonderim yapilacak satirlari secin.", vbExclamation
        Exit Sub
    End If

    Set ol = OutlookAl()
    If ol Is Nothing Then
        MsgBox "Outlook baslatilamadi.", vbCritical
        Exit Sub
    End If

    Set ws = ActiveSheet
    On Error GoTo Hata
    For Each satir In Selection.Rows
        If satir.Row > 1 Then   ' baslik satirini atla
            If EpostaGecerliMi(CStr(ws.Cells(satir.Row, "B").Value)) Then
                Set posta = ol.CreateItem(0)   ' 0 = olMailItem
                With posta
                    .To = Trim$(CStr(ws.Cells(satir.Row, "B").Value))
                    .Subject = CStr(ws.Cells(satir.Row, "C").Value)
                    .Body = "Merhaba " & ws.Cells(satir.Row, "A").Value & "," & vbCrLf & vbCrLf & _
                            ws.Cells(satir.Row, "D").Value & vbCrLf & vbCrLf & "Iyi calismalar."
                    .Display   ' taslagi ekranda ac; dogrudan gondermek icin .Send kullanin
                End With
                adet = adet + 1
            Else
                atlanan = atlanan + 1
            End If
        End If
    Next satir

    MsgBox adet & " taslak hazirlandi, gecersiz adres nedeniyle " & atlanan & " satir atlandi.", vbInformation
    Exit Sub
Hata:
    MsgBox "E-posta hazirlanamadi: " & Err.Description, vbCritical
End Sub

' Bu calisma kitabinin kaydedilmis halini ek olarak e-postaya ekler ve taslak acar
Public Sub AktifKitabiEkOlarakGonder()
    Dim ol As Object
    Dim posta As Object
    Dim aliciAdresi As String

    If ThisWorkbook.Path = "" Then
        MsgBox "Once calisma kitabini kaydedin.", vbExclamation
        Exit Sub
    End If

    aliciAdresi = InputBox("Alici e-posta adresi:", "E-posta")
    If Not EpostaGecerliMi(aliciAdresi) Then
        MsgBox "Gecerli bir e-posta adresi girilmedi.", vbExclamation
        Exit Sub
    End If

    ThisWorkbook.Save
    Set ol = OutlookAl()
    If ol Is Nothing Then Exit Sub

    On Error GoTo Hata
    Set posta = ol.CreateItem(0)
    With posta
        .To = Trim$(aliciAdresi)
        .Subject = "Rapor: " & ThisWorkbook.Name
        .Body = "Merhaba," & vbCrLf & vbCrLf & "Guncel dosya ektedir." & vbCrLf & vbCrLf & "Iyi calismalar."
        .Attachments.Add ThisWorkbook.FullName
        .Display
    End With
    Exit Sub
Hata:
    MsgBox "Hata: " & Err.Description, vbCritical
End Sub

' Secili aralik HTML tablo olarak e-posta govdesine yerlestirilir (taslak olarak acilir)
Public Sub SeciliAraligiTabloOlarakGonder()
    Dim ol As Object
    Dim posta As Object
    Dim hucre As Range
    Dim html As String
    Dim satir As Long, sutun As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 2000 Then
        MsgBox "Secim cok buyuk (en fazla 2000 hucre).", vbExclamation
        Exit Sub
    End If

    html = "<table border='1' cellpadding='4' style='border-collapse:collapse;font-family:Calibri'>"
    For satir = 1 To Selection.Rows.Count
        html = html & "<tr>"
        For sutun = 1 To Selection.Columns.Count
            Set hucre = Selection.Cells(satir, sutun)
            html = html & "<td>" & HtmlKacis(hucre.Text) & "</td>"
        Next sutun
        html = html & "</tr>"
    Next satir
    html = html & "</table>"

    Set ol = OutlookAl()
    If ol Is Nothing Then Exit Sub
    Set posta = ol.CreateItem(0)
    posta.Subject = "Tablo"
    posta.HTMLBody = "<p>Merhaba,</p>" & html & "<p>Iyi calismalar.</p>"
    posta.Display
End Sub

Private Function HtmlKacis(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlKacis = s
End Function
