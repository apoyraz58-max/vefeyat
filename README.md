# Vefeyât

**Windows deneme durumu:** Akıllı Uygulama Denetimi açık bir kullanıcı bilgisayarında `Kur.cmd`, kurulum başlamadan engellendi. Mevcut indirilebilir ZIP imzasız bir geliştirme paketidir; bu ortamda kullanıma hazır kabul edilmez. `.cmd` gerektirmeyen Windows kurucusu ve imzalı yayın üretim aracı hazırlandı; kurucu .NET Framework 4.8 referanslarıyla derlendi, 101 kayıt planı kontrolü geçti. Güvenilir bir kod imzalama sertifikası/hizmeti henüz sağlanmadığı için imzalı paket yayımlanmadı. Ayrıntılar: [imzalı dağıtım](desktop/KULLANIM-IMZALI.md).

Windows masaüstü Microsoft Word için tamamen çevrimdışı şahıs arama ve vefat tarihi ekleme eklentisi. Bu revizyon `vefeyat-ilk-surum` dalını temel alır. Word şeridinde **Vefeyât** sekmesi bulunur.

- Belgedeki şahıs adını seçin, **Vefeyât > Seçili adı ara** komutunu kullanın, doğru şahsı ve tarih ifadesini kontrol ederek ekleyin. Ad ve biçimi korunur; Word’ün Geri Al komutu kullanılabilir.
- Türkçe ve Arapça adlarda yerel arama; tarih, tanıtım ve kaynak adresi gösterimi.
- Kişisel kayıt ekleme, düzenleme, silme; CSV içe/dışa aktarma; XML yedekleme ve geri yükleme.
- Veriler `%LOCALAPPDATA%\Vefeyat\Data\catalog.xml` dosyasındadır. Yazımlar atomiktir; bir önceki sürüm `catalog.previous.xml` dosyasında tutulur. Başka Word sürecinin değişiklikleri sessizce ezilmez.
- Kurulum ve kullanım için internet, sunucu, alan adı, Docker, HTTPS sertifikası, Office.js veya VSTO gerekmez. Kaynak bağlantıları metindir; uygulama onları indirmez.

## Windows’ta kurulum

Paket ZIP dosyasını bilgisayarınızdaki bir klasöre çıkarın, Word’ü kapatın ve **Kur.cmd** dosyasını çalıştırın. Ardından Word’ü açın. Windows 10/11, .NET Framework 4.8 ve Windows masaüstü Word gerekir; 32 ve 64 bit Word için kayıt oluşturulur. Kurulum kullanıcı hesabına yapılır ve Windows’un yerleşik derleyicisini kullanır; yönetici yetkisi veya ek indirme gerektirmez.

Tam kullanım ve kaldırma yönergeleri: [desktop/KULLANIM.md](desktop/KULLANIM.md). Kaldırma kişisel verileri korur. Kurumun imzasız COM eklentisi politikasının engellediği bilgisayarlarda BT biriminizin onaylı kurulum yöntemini kullanın.

## Paket oluşturma

Python 3.10+ ve standart kütüphane yeterlidir:

```sh
python3 scripts/build-offline-package.py --csv /tam/yol/sahislar.csv --output work/Vefeyat-Windows.zip
```

CSV UTF-8 (isteğe bağlı BOM), noktalı virgülle ayrılmış şu altı sütunu içermelidir: `Kısa Ad`, `Tam Ad`, `Orijinal Ad`, `Vefat Tarihi`, `Tanıtım`, `Kaynak Link`. Gerçek kullanıcı verisi Git’e kopyalanmadan doğrudan pakete alınır. `--csv` verilmezse üç açıkça hayalî örnek kayıt paketlenir. Tarih ifadelerinde takvim dönüşümü yapılmaz.

Paket yalnız `desktop/` içindeki yerel eklenti kaynaklarını, kurulum araçlarını, kullanım kılavuzunu ve başlangıç CSV’sini içerir. Başlangıç CSV’si yalnız ilk kullanımda yerel veritabanı oluşturulurken okunur; onarım/yeniden kurulum kişisel düzenlemeleri değiştirmez.

## Geliştirme ve doğrulama

Eklenti .NET Framework 4 için C# 5 ile yazılmıştır. Office PIA ve NuGet bağımlılığı yoktur. `desktop/Connect.cs`, Office’in COM eklenti ve Ribbon arayüzlerini uygular; `CatalogForm.cs` yerel Windows Forms arayüzünü, `Core.cs` verileri, `SelectionRules.cs` belge seçimi kurallarını yönetir.

```sh
python3 -m unittest discover -s scripts/tests -p test_offline_package.py
```

Windows üzerinde yerleşik araçlarla C# testlerini çalıştırmak için:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File desktop/Test.ps1 -CsvFile 'C:\veriler\sahislar.csv'
```

Bu revizyonda C# 5 kaynakları .NET Framework 4.8 referanslarına karşı derlenmiş; veri ve seçim testleri Linux’ta .NET 8 üzerinde çalıştırılmıştır. Gerçek Windows Word yükleme, Ribbon, belge biçimlendirmesi ve geri alma kabul testleri bu ortamda çalıştırılamaz. Kontrol listesi ve doğrulama ayrıntıları: [docs/offline-validation.md](docs/offline-validation.md).

Önceki Django/PostgreSQL ve Office.js uygulamasının kaynakları tarihsel geliştirme için depoda tutulmuştur. Yeni Windows paketi bunları içermez veya çalıştırmaz. Eski web uygulamasına ait `manage.py`, `scripts/dev.sh`, `scripts/build-windows-package.py` ve Docker dosyaları yeni masaüstü kurulumunun parçası değildir.
