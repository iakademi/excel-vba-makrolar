# Excel VBA Makroları

![Lisans: MIT](https://img.shields.io/badge/Lisans-MIT-green)
![Excel](https://img.shields.io/badge/Excel-VBA-217346)
![Microsoft 365](https://img.shields.io/badge/Microsoft-365-0078D4)
![Dil](https://img.shields.io/badge/Dil-T%C3%BCrk%C3%A7e-blue)

Günlük Excel işlerini hızlandıran, Türkçe yorumlu VBA modülleri ve modern Excel formülleri için kısa başvuru kılavuzları.

## Hakkında

Bu depo; veri temizleme, rapor oluşturma, dosya işlemleri, Outlook ile e-posta gönderme ve özel çalışma sayfası fonksiyonları için yeniden kullanılabilir `.bas` modülleri içerir. `formuller/` klasöründe ise `XLOOKUP`, `FILTER`, `LET`, `LAMBDA` ve dinamik dizilere dair örnek odaklı özetler bulunur.

## İçindekiler

| Klasör | İçerik |
| --- | --- |
| [moduller](moduller/) | İçe aktarılabilir VBA modülleri (`.bas`) |
| [formuller](formuller/) | Modern Excel formülleri için Markdown kılavuzları |

## Başlarken

### Gereksinimler

- Microsoft Excel 2016 veya üzeri (formüller için `XLOOKUP`, `FILTER`, `LET`, `LAMBDA` Microsoft 365 / Excel 2021 ve sonrası gerektirir)
- `MailGonderme.bas` için masaüstü Outlook kurulu olmalıdır

### .bas dosyalarını içe aktarma

1. Dosyaları indirin veya depoyu klonlayın:
   ```bash
   git clone https://github.com/iakademi/excel-vba-makrolar.git
   ```
2. Excel'de çalışma kitabını açın ve `Alt + F11` ile VBA düzenleyicisini başlatın.
3. `Dosya > Dosya İçe Aktar...` (Import File) menüsünü seçin ve istediğiniz `.bas` dosyasını gösterin.
4. Makroları çalıştırmak için Excel'de `Alt + F8` tuşlarına basın, makroyu seçip **Çalıştır**'a tıklayın.
5. Dosyayı makro içeren biçimde kaydedin: **Excel Makro Etkin Çalışma Kitabı (*.xlsm)**.

Makro güvenliği için `Dosya > Seçenekler > Güvenlik Merkezi` altından yalnızca güvendiğiniz konumlara izin verin.

### Türkçe karakter notu

`.bas` dosyaları UTF-8 olarak kaydedilmiştir. VBA düzenleyicisi eski bir kod sayfası kullandığından yorumlardaki Türkçe karakterler içe aktarma sonrası bozuk görünebilir. Bu durum kodun çalışmasını etkilemez; ancak düzgün görünmesini isterseniz içe aktarmadan önce dosyayı Windows-1254'e çevirin:

```bash
iconv -f UTF-8 -t WINDOWS-1254 VeriTemizleme.bas > VeriTemizleme_1254.bas
```

## Klasör Yapısı

```text
excel-vba-makrolar/
├── moduller/
│   ├── README.md
│   ├── VeriTemizleme.bas
│   ├── RaporOlusturma.bas
│   ├── DosyaIslemleri.bas
│   ├── MailGonderme.bas
│   └── KullaniciTanimliFonksiyonlar.bas
├── formuller/
│   ├── README.md
│   ├── XLOOKUP.md
│   ├── FILTER.md
│   ├── LET.md
│   ├── LAMBDA.md
│   └── dinamik-diziler.md
├── LICENSE
└── README.md
```

## Katkıda Bulunma

1. Depoyu fork edin ve yeni bir dal açın.
2. Yeni makro eklerken `Option Explicit` kullanın, hata yönetimi (`On Error`) ekleyin ve Türkçe yorum yazın.
3. Dosya adlarını ASCII karakterlerle ve boşluksuz seçin.
4. Değişikliğinizi açıklayan bir Pull Request gönderin.

## Lisans

Bu proje [MIT Lisansı](LICENSE) ile sunulmaktadır. Makroları kendi sorumluluğunuzda kullanın; önemli verilerin yedeğini almadan çalıştırmayın.

## İ-Akademi Hakkında

İ-Akademi, İstanbul merkezli bir bilgi teknolojileri eğitim merkezidir; Microsoft/Office/Power Platform, Cisco, siber güvenlik, yazılım, PMP/Scrum eğitimleri sunar, Excel programlama ile Power BI, Power Automate ve Copilot danışmanlığı verir; Pearson VUE ve Kryterion sınav merkezi olarak hizmet verir. Daha fazla bilgi için: [https://iakademi.com](https://iakademi.com)
