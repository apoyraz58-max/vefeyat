# Windows bilgisayarında yerel deneme

Alan adı veya kiralık sunucu gerekmiyor. Bu seçenek Vefeyât'ı kendi Windows bilgisayarınızda Docker Desktop ile çalıştırır. Site adresi **https://localhost:8443** olur; Word de aynı bilgisayardaki bu adrese bağlanır. Bilgisayar ve Docker Desktop açıkken kullanılabilir. Başka bilgisayarlara veya internete yayımlanmaz.

Bu dosyalar Linux bulut ortamında hazırlanmıştır. Windows masaüstü Word, Windows sertifika deposu ve PowerShell akışı gerçek Windows bilgisayarında ayrıca doğrulanmalıdır; buradaki hazırlık bu doğrulamaların yapıldığı anlamına gelmez.

## 1. Ön koşullar

- Güncel Windows ve [Docker Desktop](https://docs.docker.com/desktop/setup/install/windows-install/) gerekir. Docker Desktop'ı **Linux containers** modunda başlatın. Kurum bilgisayarında kurulum veya betik çalıştırma kısıtlıysa kendi BT sürecinizi izleyin; kurum güvenlik politikasını kapatmayın.
- Güncel Microsoft 365 masaüstü Word ve WordApi 1.3 desteği gerekir. Ayrıntılar [Windows Word kılavuzunda](word-windows.md).
- Projeyi yerel bir klasöre açın; örneğin `C:\Vefeyat`. Excel dosyanızı da bilgisayarınızda bulundurun. Python, Node.js ve PostgreSQL'i ayrıca kurmanız gerekmez.
- İlk kurulumda Docker'ın resmi Python, PostgreSQL ve Caddy imajlarına, ayrıca Python paket kaynağına erişimi olmalıdır. Word görev bölmesi `https://appsforoffice.microsoft.com` üzerinden Office.js yükler.

GitHub'dan indirilen ZIP, Excel dosyanızı içermez. PowerShell'de proje klasörüne geçin ve ilk başlatmada bilgisayarınızdaki dosyanın konumunu verin:

```powershell
Set-Location C:\Vefeyat
.\scripts\windows-start.ps1 -DataFile "C:\Dosyalar\Şahıslar Veri Tabanı_251025.xlsx"
```

`.local/source/people.xlsx` zaten varsa sonraki başlatmalarda dosyayı yeniden belirtmeniz gerekmez:

```powershell
.\scripts\windows-start.ps1
```

Betik Excel'i `.local/source/people.xlsx` konumuna kopyalar. Kaynak dosyanızı değiştirmez; farklı içerikli mevcut kopyanın üzerine yazmaz. Rastgele veritabanı parolası ve uygulama anahtarı `.local/windows.env` içinde oluşturulur. Bu dosyayı paylaşmayın veya Git'e eklemeyin. Mevcut parolalar sonraki çalıştırmalarda korunur. Kayıtlar, tarih ifadeleri aynen korunarak aktarılır; aynı Excel yeniden çalıştırıldığında tekrar eklenmez.

İlk kurulumda imaj indirme ve derleme birkaç dakika sürebilir. Sonraki başlatmalarda aşağıdaki komut yeterlidir:

```powershell
.\scripts\windows-start.ps1
```

## 2. Yalnızca bu yerel sertifikayı inceleyip güvenilir yapma

Site, Caddy'nin bu kurulum için oluşturduğu yerel sertifika otoritesini kullanır. **Betik Windows güven deposunu değiştirmez.** Aşağıdaki açık adımları tamamlamadan tarayıcı veya Word'deki sertifika uyarısını geçerek ilerlemeyin; TLS doğrulamasını kapatmayın.

Önce yalnızca **herkese açık kök sertifikasını** dışarı alın:

```powershell
docker compose --env-file .local/windows.env -f compose.windows.yml cp caddy:/data/caddy/pki/authorities/local/root.crt .local/vefeyat-local-root.crt
docker compose --env-file .local/windows.env -f compose.windows.yml exec caddy sha256sum /data/caddy/pki/authorities/local/root.crt
Get-FileHash -Algorithm SHA256 .local/vefeyat-local-root.crt
certutil -dump .local/vefeyat-local-root.crt
```

İki SHA-256 değeri aynı olmalı. Sertifika bilgilerinin bu kurulumda üretilen Caddy yerel otoritesine ait olduğunu inceleyin. **Yalnızca kendi bilgisayarınızda sizin başlattığınız bu Docker kurulumundan çıkardığınız sertifikaya güvenin.** Kök otoriteye güvenmek, onun imzaladığı sertifikalara da güvenmek demektir. Özel anahtarlar korunan `caddy_data` Docker biriminde kalır; `root.key`, birim yedeği veya `.local/windows.env` dosyasını paylaşmayın.

Bu güven değişikliğini kabul ediyorsanız kendi Windows kullanıcınızın kök deposuna şu komutla ekleyin. Windows onay ekranı gösterirse sertifikayı tekrar inceleyin:

```powershell
Import-Certificate -FilePath .local/vefeyat-local-root.crt -CertStoreLocation Cert:\CurrentUser\Root
```

Bu komutun çıktısındaki sertifika **Thumbprint** değerini, deneme sonunda kaldırabilmek için not edin. Kurum ilkesi izin vermiyorsa devam etmek için kurum yöneticinizin desteklediği sertifika sürecini kullanın.

Sertifika kurulumundan sonra Word'ü ve tarayıcıyı tamamen kapatıp yeniden açın. Önce TLS doğrulaması açıkken bağlantıyı kontrol edin:

```powershell
Invoke-RestMethod https://localhost:8443/api/health/
Start-Process https://localhost:8443
```

Sağlık yanıtında `status: ok` ve `database: postgresql` beklenir. Sertifika hatası devam ediyorsa buradaki kök sertifikanın güvenini ve Docker/Caddy durumunu kontrol edin; uyarıyı atlamayın. Office'in kendi gömülü tarayıcısının da bu güveni kabul etmesi gerekir.

## 3. Yönetici hesabı

Editör hesabını etkileşimli oluşturun; parolayı sohbet ekranına veya dosyaya yazmanız gerekmez:

```powershell
docker compose --env-file .local/windows.env -f compose.windows.yml exec app python manage.py createsuperuser
```

Ardından **https://localhost:8443/admin/** adresinden oturum açın. Web üzerinden gönderilen öneriler onaylanmadan mevcut şahıs kayıtlarını değiştirmez.

## 4. Word manifesti ve eklentinin yüklenmesi

Manifest adresleri bu yerel kurulum için **https://localhost:8443** olarak ayarlanmıştır. Sertifika güveni tamamlandıktan sonra manifesti TLS doğrulaması açık biçimde indirin:

```powershell
Invoke-WebRequest https://localhost:8443/word/manifest.xml -OutFile .local/vefeyat-manifest.xml
```

Alternatif olarak uygulamanın manifest üretme komutunu kullanabilirsiniz. Komut var olan dosyanın üzerine yazmaz; benzersiz hedef adı üretiyoruz:

```powershell
$manifestPath = "/tmp/vefeyat-manifest-$([Guid]::NewGuid().ToString('N')).xml"
docker compose --env-file .local/windows.env -f compose.windows.yml exec app python manage.py export_word_manifest --base-url https://localhost:8443 --output $manifestPath
docker compose --env-file .local/windows.env -f compose.windows.yml cp "app:$manifestPath" .local/vefeyat-manifest.xml
```

Manifesti yalnızca bu dosyayı içeren, size ait paylaşılan bir klasöre koyun; `.local` klasörünü paylaşmayın. Ardından [Windows Word kılavuzundaki paylaşılan katalog adımlarını](word-windows.md) izleyin. Word'de bir deneme belgesi açıp `İbn Sînâ` yazın, yalnızca adı seçin ve Vefeyât görev bölmesinden doğru kaydı bularak tarihini ekleyin. Aynı kılavuzdaki kabul kontrollerini uygulayın.

## Durdurma, tekrar açma ve kayıtları koruma

Durumu ve günlükleri görmek için:

```powershell
docker compose --env-file .local/windows.env -f compose.windows.yml ps
docker compose --env-file .local/windows.env -f compose.windows.yml logs --tail 100 app caddy
```

Servisleri kapatmak için:

```powershell
docker compose --env-file .local/windows.env -f compose.windows.yml down
```

Bu komut veritabanını ve Caddy sertifikalarını içeren Docker birimlerini korur. Sonraki açılışta `.\scripts\windows-start.ps1` çalıştırın. Birimleri silen `down --volumes` seçeneğini kullanmayın; kayıtlar, editör değişiklikleri ve sertifika otoritesi silinir. `.local/windows.env` dosyasını da koruyun; var olan veritabanı birimi aynı parolaya ihtiyaç duyar.

Denemeyi bıraktığınızda kurduğunuz yerel kök sertifikayı kendi kullanıcı deponuzdan kaldırabilirsiniz. `certmgr.msc` aracında **Güvenilen Kök Sertifika Yetkilileri → Sertifikalar** bölümüne gidin; eklerken not ettiğiniz Thumbprint ile **yalnızca bu kurulumun sertifikasını** doğrulayıp kaldırın. İlgisiz kök sertifikaları kaldırmayın.
