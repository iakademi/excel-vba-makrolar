# Modüller

İçe aktarılabilir VBA modülleri. Her dosya bağımsızdır; yalnızca ihtiyacınız olanı alabilirsiniz.

| Modül | Açıklama | Öne çıkan prosedürler |
| --- | --- | --- |
| `VeriTemizleme.bas` | Seçili aralıktaki verileri temizler | `BosluklariTemizle`, `BosSatirlariSil`, `TekrarEdenleriSil`, `BuyukKucukHarfDuzenle` |
| `RaporOlusturma.bas` | Özet rapor sayfası ve grafik üretir | `OzetRaporOlustur`, `SayfaIcerikTablosu` |
| `DosyaIslemleri.bas` | Klasör tarama, birleştirme, PDF dışa aktarma | `KlasorDosyalariniListele`, `CalismaKitaplariniBirlestir`, `SayfalariPdfOlarakKaydet` |
| `MailGonderme.bas` | Outlook ile e-posta taslağı hazırlama/gönderme | `SeciliSatirlaraMailHazirla`, `AktifKitabiEkOlarakGonder` |
| `KullaniciTanimliFonksiyonlar.bas` | Çalışma sayfasında kullanılabilen özel fonksiyonlar | `TCKimlikDogrula`, `YaziyaCevir`, `IsGunuSayisi` |

## Kullanım

Prosedürlerin çoğu `Alt + F8` ile çalıştırılır. Bazıları önce bir aralık seçmenizi bekler; her modülün başındaki yorum bloğu ön koşulları belirtir. `KullaniciTanimliFonksiyonlar` içindeki fonksiyonlar ise hücrede `=TCKimlikDogrula(A2)` gibi doğrudan çağrılır.
