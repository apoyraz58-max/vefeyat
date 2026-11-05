# Vefeyât — imzalı Windows kurulumu

Bu kılavuz, önceden derlenmiş ve güvenilir bir yayıncı sertifikasıyla imzalanmış dağıtım içindir. Böyle bir paket henüz yayımlanmadı. Kaynak kodunun derlenmesi, dosyaların imzalandığı veya Windows tarafından kabul edildiği anlamına gelmez.

## Kullanıcı kurulumu

Windows 10/11, Windows masaüstü Word ve .NET Framework 4.8 gerekir. Word'ün 32 ve 64 bit sürümleri için kullanıcı hesabına kayıt oluşturulur; gerçek Windows kabul testleri ayrıca yapılmalıdır.

1. İmzalı paketin tamamını yerel bir klasöre çıkarın.
2. Bütün Word pencerelerini kapatın.
3. `Vefeyat.Kurulum.exe` dosyasını açıp **Kur** düğmesini kullanın. Gösterilen yayıncıyı kontrol edin.
4. Word'ü açın ve **Vefeyât** sekmesini kontrol edin.
5. Örnek verili pakette **Şahısları ara → CSV içe aktar** ile kendi bilgisayarınızdaki CSV'yi seçin. Gerçek kullanıcı verileri varsayılan dağıtım paketine eklenmez.

Kurulum, imzalı eklenti dosyasını ve başlangıç verilerini bilgisayarınıza kopyalar. Bilgisayarınızda kaynak kodu derlemez; `.cmd` veya PowerShell betiği çalıştırmaz. İmzayı yerel Windows güven bilgileriyle doğrular. Güvenilir sertifika zinciri ve gerekli Windows/Word bileşenlerinin önceden mevcut olması gerekir. Sertifika zinciri doğrulanamazsa kurulum durur; güvenlik ayarları otomatik değiştirilmez.

Çevrimdışı doğrulama imzanın bütünlüğünü ve yerel sertifika zincirini kontrol eder; sertifikanın güncel çevrimiçi iptal durumunu belirleyemez. Windows'un ve Akıllı Uygulama Denetimi'nin kendi kontrolleri ayrıca uygulanır.

Günlük arama, düzenleme, tarih ekleme, CSV aktarımı ve yedekleme tamamen çevrimdışıdır. Veriler `%LOCALAPPDATA%\Vefeyat\Data\catalog.xml` dosyasındadır. Kurulum, onarım ve kaldırma bu veri klasörünü değiştirmez.

Kaldırmak için Windows'un **Yüklü uygulamalar** listesindeki **Vefeyât** kaydını kullanın. Açılan kurucuda kaldırmayı onaylayın. Word kapalı olmalıdır. Kişisel kayıtlarınız korunur. Kaldırma aracı kendi çalışmakta olan dosyasını silemediğinde küçük kurucu dosyası yerinde kalabilir.

Akıllı Uygulama Denetimi veya kurum politikası bir dosyayı engellerse kurulumu durdurun ve uyarıyı yayıncıya iletin. İmzalama tek başına bütün kurum politikalarında veya Windows sürümlerinde kabul garantisi değildir.

## Geliştirici: imzalı dağıtım üretme

Bu aşama son kullanıcının bilgisayarında yapılmaz. Windows derleme ortamında .NET Framework derleyicisi, Windows SDK `signtool.exe`, güvenilir bir sağlayıcıdan alınmış geçerli kod imzalama sertifikası ve ilgili özel anahtara erişim gerekir. Sertifika `CurrentUser\My` deposunda bulunmalıdır. Sertifika, anahtar veya parola Git'e ya da sohbete yazılmaz.

`Build-SignedRelease.ps1`, eklentiyi ve kurucuyu önceden derler, ikisini de Authenticode ile imzalar ve imzalarını doğrular. İmzalama ve zaman damgası aşaması ağ gerektirebilir; bu gereksinim son kullanıcının çevrimdışı kurulumuna taşınmaz. Sertifika veya doğrulama eksikse dağıtım ZIP'i üretilmez.

Örnek komutta parmak izi ve zaman damgası adresini kendi imzalama altyapınıza göre doldurun:

```powershell
.\desktop\Build-SignedRelease.ps1 `
  -SigningCertificateThumbprint '<kod-imzalama-sertifikasinin-parmak-izi>' `
  -TimestampUrl '<imzalama-saglayicinizin-HTTPS-zaman-damgasi-adresi>' `
  -OutDirectory '.\work\signed-release'
```

Varsayılan paket üç hayalî örnek içerir. Açıkça yetkilendirilmiş özel bir dağıtım için `-CsvFile` verilebilir; bu durumda gerçek CSV ZIP'e girer ve dağıtım yeri ayrıca değerlendirilmelidir. Tamamlanmış paket yalnız kurucu EXE, eklenti DLL, başlangıç CSV'si, bu kılavuz ve paket bilgisi/bütünlük dosyalarını içerir.

Üretimden sonra hem EXE'nin hem DLL'nin geçerli yayıncı ve zaman damgası bilgileri, Akıllı Uygulama Denetimi açık gerçek Windows'ta kurulum, Word sekmesi, belgeye ekleme, Ctrl+Z, çevrimdışı yeniden başlatma, onarım ve kaldırma doğrulanmalıdır.

Microsoft kaynakları: [Smart App Control için kod imzalama](https://learn.microsoft.com/en-us/windows/apps/develop/smart-app-control/code-signing-for-smart-app-control), [Smart App Control sık sorulan sorular](https://support.microsoft.com/en-gb/windows/security/threat-malware-protection/smart-app-control-frequently-asked-questions).
