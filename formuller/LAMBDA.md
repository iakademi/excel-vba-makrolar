# LAMBDA

`LAMBDA`, VBA yazmadan kendi fonksiyonlarınızı tanımlamanıza izin verir. Ad Yöneticisi'ne kaydedilen bir LAMBDA, çalışma kitabında normal bir fonksiyon gibi çağrılır.

## Sözdizimi

```excel
=LAMBDA(parametre1; [parametre2; ...]; hesaplama)
```

Hücrede doğrudan test etmek için sonuna parantez içinde değer verin:

```excel
=LAMBDA(x; x*2)(5)
```

Sonuç: `10`

## Ad Yöneticisi ile Kalıcı Fonksiyon Oluşturma

1. `Formüller > Ad Yöneticisi > Yeni` yolunu izleyin.
2. **Ad** alanına fonksiyon adını yazın, örneğin `KDVLI`.
3. **Başvuru** alanına LAMBDA formülünü yazın:
   ```excel
   =LAMBDA(net; [oran]; net*(1+IF(ISOMITTED(oran); 0,20; oran)))
   ```
4. Hücrede kullanın: `=KDVLI(100)` veya `=KDVLI(100; 0,10)`.

Köşeli parantez, isteğe bağlı parametreyi belirtir; `ISOMITTED` ile kontrol edilir.

## Örnekler

### Yüzde değişim

Ad: `YUZDEDEGISIM`

```excel
=LAMBDA(eski; yeni; IF(eski=0; ""; (yeni-eski)/eski))
```

Kullanım: `=YUZDEDEGISIM(B2; C2)`

### Soyadını getir

Ad: `SOYAD`

```excel
=LAMBDA(adsoyad;
  LET(
    parcalar; TEXTSPLIT(TRIM(adsoyad); " ");
    INDEX(parcalar; COLUMNS(parcalar))
  )
)
```

Kullanım: `=SOYAD("Ayşe Nur Yılmaz")` sonucu `Yılmaz`.

### Metni baş harfleriyle kısaltma

Ad: `BASHARFLER`

```excel
=LAMBDA(metin;
  CONCAT(LEFT(TEXTSPLIT(TRIM(metin); " "); 1))
)
```

### Yardımcı fonksiyonlarla birlikte: MAP, BYROW, REDUCE

LAMBDA'nın asıl gücü dizi yardımcılarıyla ortaya çıkar.

**MAP** - her elemana uygula:

```excel
=MAP(B2:B10; LAMBDA(x; IF(x>=50; "Geçti"; "Kaldı")))
```

**BYROW** - her satır için tek değer üret (satır başı maksimum):

```excel
=BYROW(B2:E10; LAMBDA(satir; MAX(satir)))
```

**BYCOL** - her sütun için tek değer üret:

```excel
=BYCOL(B2:E10; LAMBDA(sutun; AVERAGE(sutun)))
```

**REDUCE** - değerleri tek sonuca indir (koşullu toplam):

```excel
=REDUCE(0; B2:B10; LAMBDA(toplam; x; IF(x>100; toplam+x; toplam)))
```

**SCAN** - kümülatif toplam:

```excel
=SCAN(0; B2:B10; LAMBDA(toplam; x; toplam+x))
```

### Özyinelemeli (recursive) LAMBDA

Ad Yöneticisi'nde `FAKTORIYEL` adıyla kaydedilmiş LAMBDA kendi adını çağırabilir:

```excel
=LAMBDA(n; IF(n<=1; 1; n*FAKTORIYEL(n-1)))
```

Not: Özyineleme yalnızca adı kaydedilmiş LAMBDA'larda çalışır; çağrı derinliği sınırlıdır.

## İpuçları

- Karmaşık mantığı önce hücrede `LAMBDA(...)(değer)` biçiminde test edin, sonra Ad Yöneticisi'ne taşıyın.
- Fonksiyon adlarını belirgin seçin; yerleşik fonksiyonlarla çakışmamalıdır.
- Ad Yöneticisi'ne yazarken açıklama alanını doldurun; ekip arkadaşlarınız ne yaptığını görebilir.
- LAMBDA kullanan çalışma kitabı, bu özelliği desteklemeyen Excel sürümlerinde hata verir.
