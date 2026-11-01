# LET

`LET`, formül içinde ad tanımlamanızı sağlar. Böylece aynı hesabı tekrar tekrar yazmanız gerekmez; formül hem okunur hem de daha hızlı olur.

## Sözdizimi

```excel
=LET(ad1; deger1; [ad2; deger2; ...]; hesaplama)
```

Son bağımsız değişken her zaman sonucu döndüren hesaplamadır.

## Örnekler

### Tekrarı önleme

LET olmadan:

```excel
=IF(SUM(B2:B50)/COUNT(B2:B50)>100; SUM(B2:B50)/COUNT(B2:B50)*1.1; SUM(B2:B50)/COUNT(B2:B50))
```

LET ile:

```excel
=LET(
  ort; AVERAGE(B2:B50);
  IF(ort>100; ort*1.1; ort)
)
```

### KDV hesabı

```excel
=LET(
  net; B2;
  oran; 0,20;
  kdv; net*oran;
  net+kdv
)
```

### Ad tanımlayıp FILTER ile birlikte kullanma

Ortalamanın üzerindeki satışlar:

```excel
=LET(
  veri; Satis[Tutar];
  ort; AVERAGE(veri);
  FILTER(Satis; veri>ort; "Yok")
)
```

### Metin ayrıştırma

`"Ad Soyad"` biçimindeki hücreden soyadını almak:

```excel
=LET(
  metin; TRIM(A2);
  bosluk; FIND(" "; metin & " ");
  ad; LEFT(metin; bosluk-1);
  soyad; MID(metin; bosluk+1; 100);
  soyad & ", " & ad
)
```

### Yüzde değişim ve güvenli bölme

```excel
=LET(
  onceki; B2;
  simdiki; C2;
  IF(onceki=0; "-"; (simdiki-onceki)/onceki)
)
```

## İpuçları

- Ad tanımları yukarıdan aşağıya hesaplanır; bir ad yalnızca kendisinden önce tanımlanan adları kullanabilir.
- Adlar harfle başlamalı, hücre adresi gibi (`A1`, `R1C1`) görünmemelidir.
- Satır sonları ve girinti ekleyerek (`Alt + Enter`) uzun formülleri okunur kılın.
- `LET` ile tanımlanan ad yalnızca o formülün içinde geçerlidir; tekrar kullanmak için `LAMBDA` ve Ad Yöneticisi gerekir.
