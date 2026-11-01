Attribute VB_Name = "KullaniciTanimliFonksiyonlar"
'==============================================================
' Modul   : KullaniciTanimliFonksiyonlar
' Amac    : Calisma sayfasinda dogrudan kullanilabilen ozel fonksiyonlar (UDF)
' Ornek   : =TCKimlikDogrula(A2)   =YaziyaCevir(1234,5)   =IsGunuSayisi(A2;B2)
'==============================================================
Option Explicit

' T.C. kimlik numarasi algoritmik dogrulamasi (kisinin gercek olup olmadigini sorgulamaz)
' Kurallar: 11 hane, ilk hane 0 degil,
'   10. hane = ((1,3,5,7,9. hanelerin toplami)*7 - (2,4,6,8. hanelerin toplami)) mod 10
'   11. hane = ilk 10 hanenin toplami mod 10
Public Function TCKimlikDogrula(ByVal deger As Variant) As Boolean
    Dim s As String
    Dim d(1 To 11) As Long
    Dim i As Long
    Dim tekler As Long, ciftler As Long, toplam As Long

    s = Trim$(CStr(deger))
    If Len(s) <> 11 Or Not s Like String(11, "#") Then Exit Function
    If Left$(s, 1) = "0" Then Exit Function

    For i = 1 To 11
        d(i) = CLng(Mid$(s, i, 1))
    Next i

    tekler = d(1) + d(3) + d(5) + d(7) + d(9)
    ciftler = d(2) + d(4) + d(6) + d(8)
    If ((tekler * 7 - ciftler) Mod 10 + 10) Mod 10 <> d(10) Then Exit Function

    For i = 1 To 10
        toplam = toplam + d(i)
    Next i
    TCKimlikDogrula = (toplam Mod 10 = d(11))
End Function

' Iki tarih arasindaki is gunu sayisi (hafta sonlari haric; istege bagli tatil listesi)
Public Function IsGunuSayisi(ByVal baslangic As Date, ByVal bitis As Date, _
                             Optional ByVal tatiller As Range) As Long
    Dim gun As Date
    Dim say As Long
    Dim t As Range
    Dim tatilMi As Boolean

    If bitis < baslangic Then Exit Function
    For gun = baslangic To bitis
        If Weekday(gun, vbMonday) <= 5 Then
            tatilMi = False
            If Not tatiller Is Nothing Then
                For Each t In tatiller.Cells
                    If IsDate(t.Value) Then
                        If CDate(t.Value) = gun Then tatilMi = True: Exit For
                    End If
                Next t
            End If
            If Not tatilMi Then say = say + 1
        End If
    Next gun
    IsGunuSayisi = say
End Function

' Sayiyi Turkce yaziya cevirir (ornek: 1234,5 -> "Bin Iki Yuz Otuz Dort Lira Elli Kurus")
Public Function YaziyaCevir(ByVal sayi As Double) As String
    Dim lira As Double
    Dim kurus As Long
    Dim sonuc As String

    If sayi < 0 Or sayi >= 1000000000000# Then
        YaziyaCevir = "#Aralik disi"
        Exit Function
    End If

    lira = Int(sayi)
    kurus = CLng(Round((sayi - lira) * 100, 0))
    If kurus = 100 Then lira = lira + 1: kurus = 0

    If lira = 0 And kurus = 0 Then
        YaziyaCevir = "Sifir Lira"
        Exit Function
    End If

    If lira > 0 Then sonuc = SayiYazi(lira) & " Lira"
    If kurus > 0 Then sonuc = Trim$(sonuc & " " & SayiYazi(kurus) & " Kurus")
    YaziyaCevir = sonuc
End Function

Private Function SayiYazi(ByVal n As Double) As String
    Dim birler As Variant, onlar As Variant, basamak As Variant
    Dim grup As Long, kalan As Long, sonuc As String, i As Long
    Dim parca As String

    birler = Array("", "Bir", "Iki", "Uc", "Dort", "Bes", "Alti", "Yedi", "Sekiz", "Dokuz")
    onlar = Array("", "On", "Yirmi", "Otuz", "Kirk", "Elli", "Altmis", "Yetmis", "Seksen", "Doksan")
    basamak = Array("", "Bin", "Milyon", "Milyar")

    If n = 0 Then SayiYazi = "Sifir": Exit Function

    i = 0
    Do While n > 0
        grup = CLng(n - Int(n / 1000) * 1000)
        n = Int(n / 1000)
        If grup > 0 Then
            parca = UcHane(grup, birler, onlar)
            ' "Bir Bin" denmez, dogrudan "Bin" denir
            If i = 1 And grup = 1 Then parca = ""
            sonuc = Trim$(parca & " " & basamak(i) & " " & sonuc)
        End If
        i = i + 1
    Loop
    SayiYazi = sonuc
End Function

Private Function UcHane(ByVal n As Long, birler As Variant, onlar As Variant) As String
    Dim y As Long, o As Long, b As Long
    y = n \ 100: o = (n Mod 100) \ 10: b = n Mod 10
    If y = 1 Then
        UcHane = "Yuz"
    ElseIf y > 1 Then
        UcHane = birler(y) & " Yuz"
    End If
    UcHane = Trim$(UcHane & " " & onlar(o) & " " & birler(b))
End Function

' Hucredeki metinden yalnizca rakamlari dondurur (ornek: "Tel: 0 (212) 555" -> "0212555")
Public Function SadeceRakamlar(ByVal metin As String) As String
    Dim i As Long, c As String, sonuc As String
    For i = 1 To Len(metin)
        c = Mid$(metin, i, 1)
        If c Like "#" Then sonuc = sonuc & c
    Next i
    SadeceRakamlar = sonuc
End Function

' Verilen aralikta, belirtilen metne esit hucrelerin yazi rengi/arka plan rengine
' gore sayisini verir. Kullanim: =RenkSay(A1:A20; D1)  (D1 = ornek renkli hucre)
Public Function RenkSay(ByVal aralik As Range, ByVal ornek As Range) As Long
    Dim h As Range
    Application.Volatile True   ' renk degisimi hesaplamayi tetiklemez; F9 ile yenileyin
    For Each h In aralik.Cells
        If h.Interior.Color = ornek.Interior.Color Then RenkSay = RenkSay + 1
    Next h
End Function
