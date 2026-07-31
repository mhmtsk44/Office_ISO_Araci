# Office Taşınabilir ISO Aracı (v5.0.9)

Microsoft'un resmi **Office Deployment Tool (ODT)**'unu kullanarak, seçtiğiniz Office paketi/sürümü/dile göre kurulum dosyalarını indirip **ISO kalıbı** haline getiren Windows GUI aracı.

**Hazırlayan:** Mehmet IŞIK

## Ne Yapar?

1. Paket (Temel/Özel uygulamalar), sürüm (365 / LTSC 2024), mimari (x86/x64) ve dil seçimini adım adım alır.
2. Microsoft'un resmi ODT'sini indirir, seçimlere göre bir yapılandırma XML'i oluşturur.
3. `setup.exe /download` ile Office kurulum paketlerini **doğrudan Microsoft sunucularından** indirir.
4. İndirilen dosyalardan IMAPI2 (veya yedek olarak `oscdimg`) ile bir `.iso` dosyası oluşturur, SHA256 bütünlük dosyası ekler.
5. ISO motoru çalışmazsa dosyaları klasör olarak hazırlar (USB'ye kopyalanabilir).

## ⚠️ Lisans ve Etkinleştirme Hakkında

Bu araç **yalnızca kurulum ortamı (ISO) hazırlar** — indirilen dosyalar Microsoft'un resmi ODT/CDN kaynağından gelir.

- Bu script içinde **etkinleştirme, kırma (crack), KMS emülatörü veya lisans anahtarı üretme gibi hiçbir mekanizma yoktur.**
- Oluşturulan ISO ile Office'i kurduktan sonra, kullanabilmek için **geçerli ve satın alınmış bir Microsoft 365 aboneliği veya Office LTSC lisansı/ürün anahtarı** ile etkinleştirme yapmanız gerekir.
- LTSC 2024 seçiminde araç bunu arayüzde de belirtir: *"ISO yalnızca kurulum kaynağıdır, lisans içermez."*
- Lisanssız/etkinleştirilmemiş Office kısıtlı modda (bildirim şeridi, düzenleme kapalı vb.) çalışır; bu bir araç kısıtlaması değil, Microsoft'un lisans politikasıdır.

Lisans satın almak için: [Microsoft 365 resmi sitesi](https://www.microsoft.com/microsoft-365)

## Gereksinimler

- Windows + PowerShell (yönetici yetkisi, otomatik UAC yükseltme ister)
- İnternet bağlantısı (indirme Microsoft sunucularından yapılır)
- Yeterli disk alanı (~2-6 GB, seçilen uygulama sayısına göre)

## Kullanım

1. Scripti indirin, sağ tık → **PowerShell ile Çalıştır**.
2. UAC isteğini onaylayın.
3. Sihirbazda paket/sürüm/mimari/dil/hedef klasör seçimlerini yapın.
4. İndirme ve ISO oluşturma tamamlanınca dosya hedef klasörde hazır olur (yanında `.sha256` dosyasıyla).

```powershell
iwr "https://raw.githubusercontent.com/mhmtsk44/Office_ISO_Araci/refs/heads/main/Office_ISO_Araci_v5_0_DPI_Aware.ps1" -OutFile "$env:TEMP\Office_ISO.ps1"
powershell -ExecutionPolicy Bypass -File "$env:TEMP\Office_ISO.ps1""
```

> ⚠️ Kaynağı doğrulamadan internetten indirip çalıştırdığınız her script gibi dikkatli olun; araç yönetici yetkisiyle çalışır ve sisteme dosya indirir/yazar.

## Not

Geçici indirme/çıkartma işlemleri `%TEMP%` altında benzersiz bir klasörde yapılır, izin sorunlarını önler.

## Lisans (Bu Aracın Kendisi)

Bu araç serbestçe kullanılabilir, değiştirilebilir ve dağıtılabilir. Kaynak belirtmek zorunlu değil, takdir edilir.

*(Karışıklığı önlemek için: yukarıdaki "Lisans" başlığı bu PowerShell scriptinin kullanım koşuluyla ilgilidir; Office ürününün lisansıyla bir ilgisi yoktur.)*
