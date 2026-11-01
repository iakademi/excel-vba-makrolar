# FILTER

`FILTER`, bir aralıktaki satırları veya sütunları koşula göre süzer; sonuç dinamik olarak taşar (spill).

## Sözdizimi

```excel
=FILTER(dizi; dahil_et; [bos_ise])
```

| Bağımsız değişken | Açıklama |
| --- | --- |
| `dizi` | Süzülecek aralık |
| `dahil_et` | Satır sayısı kadar `DOĞRU`/`YANLIŞ` döndüren koşul |
| `bos_ise` | Hiç kayıt yoksa gösterilecek değer (verilmezse `#HESAPLA!`) |

## Örnekler

### Tek koşul

İstanbul'daki müşteriler:

```excel
=FILTER(Musteriler; Musteriler[Sehir]="İstanbul"; "Kayıt yok")
```

### VE koşulu (çarpma)

Satışı 10.000'den büyük **ve** bölgesi "Marmara" olanlar:

```excel
=FILTER(Satis; (Satis[Tutar]>10000)*(Satis[Bolge]="Marmara"); "Kayıt yok")
```

### VEYA koşulu (toplama)

Durumu "Beklemede" **veya** "Gecikmiş" olanlar:

```excel
=FILTER(Gorevler; (Gorevler[Durum]="Beklemede")+(Gorevler[Durum]="Gecikmiş"); "Kayıt yok")
```

### Belirli sütunları döndürme

`CHOOSECOLS` ile yalnızca 1. ve 3. sütunu al:

```excel
=CHOOSECOLS(FILTER(Satis; Satis[Bolge]="Ege"); 1; 3)
```

### Metin içeren kayıtlar

Adında "Power" geçen kurslar (`SEARCH` büyük/küçük harf duyarsızdır):

```excel
=FILTER(Kurslar; ISNUMBER(SEARCH("Power"; Kurslar[Ad])); "Eşleşme yok")
```

### Açılır listeyle birleştirme

`G1` hücresindeki seçime göre listeyi sıralı göster:

```excel
=SORT(FILTER(Satis; Satis[Urun]=G1; "Kayıt yok"); 3; -1)
```

### Boş olmayan satırlar

```excel
=FILTER(A2:D100; A2:A100<>"")
```

### Sonuç sayısı

```excel
=ROWS(FILTER(Satis; Satis[Bolge]="Ege"))
```

## Dikkat Edilecekler

- Sonucun taşacağı hücreler boş olmalıdır; aksi halde `#TAŞMA!` hatası alırsınız.
- `dahil_et` boyutu `dizi` ile uyumlu olmalıdır (satır filtresi için sütun vektörü).
- Tablo adları kullanırsanız veri eklendikçe sonuç otomatik güncellenir.
