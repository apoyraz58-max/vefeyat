# Vefeyât ilk deneme — kurulum engeli

**Bilinen durum:** Kullanıcıdaki Akıllı Uygulama Denetimi, `Kur.cmd` dosyası açılırken kurulumu engelledi. Bu dalda daha önce yayımlanan `Vefeyat-Windows-Kurulum.zip`, imzasız geliştirme paketidir. Bu koruma açık bilgisayarlarda kurulum için önerilmez; aynı ZIP'i yeniden indirmek bu dağıtım eksikliğini düzeltmez.

`.cmd` gerektirmeyen kurucu kaynakları `desktop/Installer.cs` dosyasında, imzalı yayın üretimi `desktop/Build-SignedRelease.ps1` dosyasında hazırlandı. Son kullanıcıya verilecek yeni paket önceden derlenmiş ve güvenilir yayıncı tarafından imzalanmış kurucu EXE ve eklenti DLL içermelidir. Gerekli imzalama sertifikası/hizmeti sağlanmadığından henüz imzalı paket yayımlanmadı. Gerçek Windows Word/SAC kabul testi bekliyor.

Ayrıntılar: [İmzalı kurulum ve üretim gereksinimleri](../desktop/KULLANIM-IMZALI.md).

Mevcut örnek ZIP yalnız üç hayalî kayıt içerir; kişisel CSV GitHub'a yüklenmedi. Veriler uygulama kullanılırken bilgisayarda tutulur.
