# XLOOKUP

`XLOOKUP`, `VLOOKUP` (DÜŞEYARA) ve `HLOOKUP` fonksiyonlarının yerini alan, her yöne arama yapabilen fonksiyondur.

## Sözdizimi

```excel
=XLOOKUP(aranan_deger; arama_dizisi; donus_dizisi; [bulunamazsa]; [eslesme_modu]; [arama_modu])
```

| Bağımsız değişken | Açıklama |
| --- | --- |
| `aranan_deger` | Aranan değer |
| `arama_dizisi` | Aramanın yapılacağı tek sütun veya satır |
| `donus_dizisi` | Sonucun alınacağı aralık |
| `bulunamazsa` | Eşleşme yoksa gösterilecek değer (`#YOK` yerine) |
| `eslesme_modu` | `0` tam eşleşme (varsayılan), `-1` tam veya bir küçük, `1` tam veya bir büyük, `2` joker karakter |
| `arama_modu` | `1` baştan sona (varsayılan), `-1` sondan başa, `2`/`-2` ikili arama |

## Örnekler

### Temel arama

Ürün kodundan fiyatı getir:

```excel
=XLOOKUP(A2; Urunler[Kod]; Urunler[Fiyat]; "Bulunamadı")
```

### Soldaki sütunu döndürme

`VLOOKUP` yalnızca sağa bakabilir; `XLOOKUP` her iki yönde çalışır:

```excel
=XLOOKUP(E2; C2:C100; A2:A100)
```

### Birden fazla sütun döndürme

Dönüş aralığı birden çok sütun içerirse sonuç yana doğru taşar:

```excel
=XLOOKUP(A2; Musteriler[ID]; Musteriler[[Ad]:[Sehir]])
```

### Sondan başa arama (son kayıt)

Bir müşterinin en son siparişini bulmak için:

```excel
=XLOOKUP(A2; Siparisler[Musteri]; Siparisler[Tutar]; ; 0; -1)
```

### Yaklaşık eşleşme (kademeli oran)

Satış tutarına göre komisyon oranını bulma (alt sınırlar artan sırada olmak zorunda değildir):

```excel
=XLOOKUP(B2; Oranlar[AltSinir]; Oranlar[Oran]; 0; -1)
```

### Joker karakterle arama

"Excel" ile başlayan ilk kursu bulma:

```excel
=XLOOKUP("Excel*"; Kurslar[Ad]; Kurslar[Ad]; "Yok"; 2)
```

### İki ölçütle arama

Ölçütleri `&` ile birleştirin:

```excel
=XLOOKUP(A2&"|"&B2; Veri[Bolge]&"|"&Veri[Urun]; Veri[Satis]; "Yok")
```

### İki yönlü arama (satır + sütun kesişimi)

```excel
=XLOOKUP(G1; A2:A10; XLOOKUP(G2; B1:F1; B2:F10))
```

## Dikkat Edilecekler

- Arama ve dönüş dizilerinin boyutu aynı olmalıdır.
- Tam eşleşme varsayılandır; `VLOOKUP`'taki gibi `YANLIŞ` yazmaya gerek yoktur.
- Büyük/küçük harf duyarlı değildir; duyarlı arama için `EXACT` ile `FILTER` kullanın.
