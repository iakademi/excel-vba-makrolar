# Dinamik Diziler

Microsoft 365 ve Excel 2021 ile birlikte bir formül birden fazla değer döndürebilir. Sonuç, formülün bulunduğu hücreden başlayarak komşu hücrelere **taşar** (spill). Eski sürümlerdeki `Ctrl + Shift + Enter` ile girilen dizi formüllerine artık gerek yoktur.

## Taşma Aralığı

Taşan sonucun tamamına `#` işaretiyle başvurulur:

```excel
=SUM(A2#)
```

`A2` hücresindeki formülün tüm taşma aralığını toplar. Aralık büyüyüp küçüldükçe formül otomatik uyum sağlar.

## Temel Dizi Fonksiyonları

| Fonksiyon | Amaç | Örnek |
| --- | --- | --- |
| `UNIQUE` | Benzersiz değerler | `=UNIQUE(Satis[Bolge])` |
| `SORT` | Sıralama | `=SORT(A2:C20; 3; -1)` |
| `SORTBY` | Başka sütuna göre sıralama | `=SORTBY(A2:A20; C2:C20; -1)` |
| `FILTER` | Koşullu süzme | `=FILTER(A2:C20; C2:C20>100)` |
| `SEQUENCE` | Sayı dizisi | `=SEQUENCE(5; 1; 1; 1)` |
| `RANDARRAY` | Rastgele sayılar | `=RANDARRAY(3; 2; 1; 100; TRUE)` |
| `TAKE` / `DROP` | Baştan/sondan satır alma/atma | `=TAKE(SORT(A2:B50; 2; -1); 5)` |
| `VSTACK` / `HSTACK` | Aralıkları alt alta / yan yana birleştirme | `=VSTACK(A2:B5; D2:E8)` |
| `TOCOL` / `TOROW` | Aralığı tek sütuna/satıra çevirme | `=TOCOL(A2:C5; 1)` |
| `TEXTSPLIT` | Metni ayırıcıya göre bölme | `=TEXTSPLIT(A2; ", ")` |

## Örnek Senaryolar

### Benzersiz liste ve toplamları

Önce bölgelerin benzersiz listesi `E2` hücresine:

```excel
=SORT(UNIQUE(Satis[Bolge]))
```

Yanındaki `F2` hücresine her bölgenin toplamı (taşan listeyi `#` ile kullanır):

```excel
=SUMIFS(Satis[Tutar]; Satis[Bolge]; E2#)
```

### En çok satan beş ürün

```excel
=TAKE(SORTBY(Urunler; Urunler[Satis]; -1); 5)
```

### Takvim günleri üretme

Bir ayın tüm günlerini `A1` hücresindeki ay başlangıcına göre listeleme:

```excel
=SEQUENCE(DAY(EOMONTH(A1; 0)); 1; A1; 1)
```

Hücre biçimini tarih yapmayı unutmayın.

### Çarpım tablosu

```excel
=SEQUENCE(9; 1) * SEQUENCE(1; 9)
```

### Sütunlara bölünmüş metni birleştirip tek listeye çevirme

```excel
=TOCOL(A2:C10; 1)
```

İkinci argüman `1` boş hücreleri yok sayar.

### Birden fazla tablodan tek liste

```excel
=VSTACK(Ocak[Ad]; Subat[Ad]; Mart[Ad])
```

## Sık Karşılaşılan Hatalar

| Hata | Nedeni | Çözüm |
| --- | --- | --- |
| `#TAŞMA!` | Taşma aralığında dolu hücre var | Engel olan hücreleri temizleyin |
| `#HESAPLA!` | Sonuç boş küme (ör. `FILTER` eşleşme bulamadı) | `FILTER`'ın `bos_ise` argümanını kullanın |
| `#ADI?` | Sürüm fonksiyonu desteklemiyor veya adı yanlış | Microsoft 365 sürümünüzü güncelleyin |

## İpuçları

- Taşan aralıkta tek bir hücreyi silmek yerine kaynak formülü silin.
- Tablolar (`Ctrl + T`) içinde taşma formülü kullanılamaz; formülü tablo dışına yazın.
- Büyük veri kümelerinde `UNIQUE` ve `FILTER` zincirlerini `LET` ile sadeleştirin.
