# Çevrimdışı Windows Word doğrulaması

Bu belge yerel C#/.NET Framework COM eklentisinin kabul kontrolüdür. Kaynak ve otomatik veri/seçim testleri, Word'ün COM eklentisini yüklediğini veya belgedeki biçimlendirmeyi koruduğunu tek başına kanıtlamaz. Gerçek Windows masaüstü Word testleri aşağıdaki tablolar dolduruluncaya kadar **beklemede** kabul edilmelidir. Önceki web/Office.js sürümünün testleri bu paketin Windows kabulü yerine geçmez.

## Otomatik testlerin kapsamı

`desktop/tests/CoreTests.cs`, üçüncü taraf test paketi kullanmadan `Core.cs` ile derlenen bir test uygulamasıdır. Kapsamı UTF-8 CSV, alıntılı ayraç ve çok satırlı alanlar, Arapça metinler, XML kalıcılığı, Türkçe ve işaretli karakterlerden bağımsız arama, sıralama, kayıt ekleme/düzenleme/silme, tekrarlanan aktarımlar, belirsiz adaş kayıtları, hatalı aktarımda değişiklik yapılmaması, eski veriyle yazma çakışması ve XML yedek/geri yüklemedir. Başlangıç CSV'si argüman olarak verilirse 7.934 kaydın tamamı ve yeniden aktarımda kayıt sayısının korunması da denetlenir.

`desktop/tests/SelectionTests.cs`, ad/tarih doğrulamasını, izlenen seçimin metin/konum/öykü karşılaştırmasını ve yinelenen tarih kontrollerini Word başlatmadan sınar. `desktop/tests/ComContractTests.cs`, Office COM sözleşmelerini, şerit XML'ini ve komutların callback eşleşmelerini yansıma ile kontrol eder; Office'in eklentiyi gerçekten yüklediğini sınamaz. Otomatik testlerin gerçek sonuç ve komutlarını paket üretiminin doğrulama raporuna kaydedin; bu kontrol listesinde sonuç üretilmeden “geçti” yazmayın.

## 6 Kasım 2026 otomatik doğrulama kaydı

- Dört eklenti kaynağı C# 5 dil kipiyle .NET Framework 4.8 referanslarına karşı başarıyla derlendi.
- `CoreTests`: .NET 8 üzerinde 45 kontrol geçti; sağlanan CSV'nin 7.934 kaydı, yeniden aktarımı ve araması doğrulandı.
- `SelectionTests`: .NET 8 üzerinde 27 kontrol geçti.
- `ComContractTests`: .NET Framework referanslarına karşı derlenen eklenti üzerinde 38 yansıma/XML kontrolü geçti.
- `python3 -m unittest discover -s scripts/tests -p test_offline_package.py`: 6 test geçti.
- `Install.ps1`, `Uninstall.ps1`, `Test.ps1`: PowerShell 7.4.6 ayrıştırıcısıyla sözdizimi kontrolü geçti. Windows PowerShell 5.1 üzerinde çalıştırma beklemededir.
- Gerçek CSV'yi içeren ZIP'in dosya bütünlüğü, SHA-256 manifesti, 7.934 kayıt sayısı ve kaynak CSV ile bayt eşitliği doğrulandı. Gerçek veri Git'e alınmadı.

Bu kontroller Windows COM etkinleştirmesini veya Word belge işlemini çalıştırmaz.

## Windows test kaydı

Her deneme için Windows sürümü, Word ürün adı/sürümü ve **Word bitliği** (Dosya → Hesap → Word Hakkında), standart kullanıcı hesabı, .NET Framework sürümü, paket sürümü ve test tarihini kaydedin. Windows'un bitliği ile Word'ün bitliği ayrı bilgilerdir.

| Ortam | Durum | Kanıt |
| --- | --- | --- |
| Windows 10/11, 64 bit Windows üzerinde 32 bit masaüstü Word | Beklemede | Sürüm, kurulum kaydı, Vefeyât sekmesi |
| Windows 10/11, 64 bit masaüstü Word | Beklemede | Sürüm, kurulum kaydı, Vefeyât sekmesi |
| Kurumca yönetilen cihazdaki eklenti politikası | Ortama bağlı | Varsa izin verilen dağıtım yöntemi; korumalar kapatılmaz |

## Kurulum ve çevrimdışı çalışma

- [ ] Paketi yerel, OneDrive ile eşitlenmeyen bir klasöre tamamen çıkarın. Standart kullanıcıyla, yönetici yetkisi vermeden `Kur.cmd` çalıştırın.
- [ ] Word açıkken kurulumun açıklayıcı hata vererek durduğunu doğrulayın; Word kapandıktan sonra kurulum tamamlanmalıdır.
- [ ] Ağ bağlantısını kesip uçak modunu etkinleştirin. Gerekli Windows/Word bileşenleri önceden kurulu olmalıdır. Kurulumu ve ilk katalog açılışını bağlantı olmadan tekrarlayın; indirme veya oturum açma istenmemelidir.
- [ ] Word'ü açın. Şeritte Vefeyât ve `Seçili adı ara`, `Şahısları ara`, `Veri klasörü`, `Kullanım` komutlarını doğrulayın.
- [ ] İlk açılışta 7.934 kaydın tamamının yerel veri dosyasına alındığını doğrulayın. Aramada en çok 200 satır görünmesi toplam kayıt sayısıyla karıştırılmamalıdır.
- [ ] Docker Desktop, web sunucusu veya veritabanı hizmeti çalışmadan arama ve belgeye ekleme yapın. Eklentinin bir dinleme portu, yerel sunucu veya arka plan sunucu süreci başlatmadığını süreç/port görünümüyle doğrulayın.
- [ ] Kaynak URL'lerinin yalnız metin olarak gösterildiğini, arama ve ekleme sırasında tarayıcı açılmadığını veya kaynak sitelerden veri alınmadığını doğrulayın.
- [ ] Word'ü ve bilgisayarı yeniden başlatın. Çevrimdışı arama, değiştirilmiş kayıtlar ve Vefeyât sekmesi korunmalıdır.

## Belgeye ekleme

- [ ] Düzenlenebilir belgede `Abaka` adını seçin. `Seçili adı ara` ile doğru kaydı açıkça seçip tarih önizlemesini kontrol edin. Sonuç `Abaka (ö. 680/1282)` olmalıdır.
- [ ] Aynı paragrafta iki kez geçen aynı adın yalnız seçilen örneğine tarih eklenmelidir.
- [ ] Kalın, italik, farklı yazı tipi, farklı renk ve karışık biçimlendirmeli adlarla deneme yapın. Adın metni ve biçimi korunmalı, çevredeki metin değişmemelidir.
- [ ] Adın sonundaki seçime dahil boşluk ve kesintisiz boşlukla deneme yapın. Tarih adın sonuna gelmeli; mevcut metin silinmemelidir.
- [ ] Bir eklemeden sonra Ctrl+Z ile işlemi geri alın. Ad ve çevresindeki metin ilk durumuna dönmelidir; Word'ün önceki geri alma geçmişi korunmalıdır.
- [ ] Tarihi zaten hemen sonrasında bulunan adı yeniden seçin. Aynı tarih ikinci kez eklenmemelidir.
- [ ] Adaş kayıtlarda sonuçlar otomatik olarak belgeye yazılmamalıdır; kullanıcı doğru şahsı seçmelidir.
- [ ] Boş seçim, yalnız imleç, paragraf sonu içeren seçim, çok uzun seçim, tablo hücresi sonu içeren seçim ve boş/çok satırlı tarih reddedilmelidir.
- [ ] Salt okunur, korumalı ve Korumalı Görünüm'deki belgelerde ekleme yapılmamalı, anlaşılır bildirim verilmelidir. Güvenlik korumasını test amacıyla kapatmayın.
- [ ] Arama penceresi açıkken seçim/belge değişimini test edebilen kontrollü bir Word otomasyonu kullanın. Aynı metnin başka konumu dahil değişen hedefte ekleme yapılmamalıdır. Modal pencerenin normal kullanıcı girişini engellemesi, bu kod yolunun tek başına testi sayılmaz.
- [ ] Arapça ad, orijinal Arapça alan, tablo içindeki ad ve üstbilgi/altbilgi gibi farklı belge öykülerinde hedef doğruluğunu denetleyin.
- [ ] `Şahısları ara` komutuyla açılan salt katalog akışının Word belgesine yazmadığını doğrulayın.

## Veri kalıcılığı ve yedekler

- [ ] `%LOCALAPPDATA%\Vefeyat\Data\catalog.xml` dosyasının oluşturulduğunu ve verilerin eklenti kurulum klasöründen ayrı tutulduğunu doğrulayın.
- [ ] Bir kayıt ekleyin, tarihini düzenleyin, Word'ü kapatıp yeniden açın. Düzenleme korunmalıdır. Silme yalnız onaylanan seçili kaydı etkilemelidir.
- [ ] Türkçe `İ/ı`, `â/î/û` ve Arapça karakterlerle arayın. Tarihin orijinal metin olduğunu, takvim dönüşümü yapılmadığını doğrulayın.
- [ ] Başlangıç CSV'sini iki kez içe aktarın; sayı artmamalıdır. Tekil kimlikte tarih/tanıtım değişikliği mevcut kaydı güncellemeli; belirsiz adaş değişikliği rastgele bir kaydı ezmemelidir.
- [ ] Son satırı bozuk CSV'yi içe aktarın; hiçbir satır kısmen kaydedilmemelidir. Geçersiz kodlama ve eksik başlıklar anlaşılır hata vermelidir.
- [ ] CSV dışa aktarımını tekrar açın; bütün kayıtlar, altı başlık, Unicode ve tırnaklı/çok satırlı alanlar korunmalıdır.
- [ ] XML yedek alın; değişiklik yapıp yedeği geri yükleyin. Kayıtlar ve kimlikleri yedekteki duruma dönmelidir. Bozuk XML/DTD içeren dosya reddedilmeli ve mevcut veri korunmalıdır.
- [ ] İki ayrı Word işlemi üzerinden aynı kataloğu açın. İlk işlem kaydettikten sonra eski veriye dayanan ikinci yazma reddedilmeli; pencere yeniden açıldığında güncel veriler görünmelidir.
- [ ] Önceki sürüm dosyasının başarılı değişiklikten önceki durumu tuttuğunu ve bağımsız yedeklemenin bu dönen dosyadan ayrı olduğunu doğrulayın.
- [ ] Test verisiyle, yedek alındıktan sonra bozuk yerel dosyayı açmayı deneyin. Veriler sessizce sıfırlanmamalı; açılışta sunulan XML kurtarma seçeneğiyle geçerli yedekten dönülebilmelidir. İptalde çalışma dosyası değişmemelidir. Kurtarmada değiştirilen çalışma dosyasının `catalog.previous.xml` olarak tutulduğunu denetleyin; kurtarmadan önce mevcut önceki-sürüm dosyasını ayrıca kopyalayın.

## Onarım ve kaldırma

- [ ] Kişisel veri dosyasının özetini kaydedin. Word kapalıyken `Kur.cmd` dosyasını tekrar çalıştırın; kurulum yinelenen COM kayıtları üretmemeli ve kişisel veri özeti aynı kalmalıdır.
- [ ] Word açıkken `Kaldir.cmd` dosyası durmalıdır. Word kapalıyken kaldırma sonrası sekme/kayıt kalkmalı, kişisel veri dosyası ve ayrı yedekler korunmalıdır.
- [ ] Yeniden kurulumda önceden düzenlenmiş katalog açılmalı; başlangıç CSV'si kişisel değişiklikleri ezmemelidir.
- [ ] Eski Office.js eklentisi hâlâ kuruluysa olası ikinci sekmenin eski eklentiye ait olduğunu doğrulayın. Yalnız eski eklentiyi Word'ün Eklentilerim bölümünden kaldırınca yerel sekme kalmalıdır.
- [ ] Kurumsal politika yüklemeyi engelliyorsa sonucu “politika nedeniyle engellendi” olarak kaydedin ve onaylı dağıtım isteyin; korumaları devre dışı bırakmayı kabul adımı olarak kullanmayın.

Kabul kaydına tamamlanan maddeleri, başarısız maddeleri, Word sürümlerini ve kanıtları ekleyin. Bekleyen Windows/Word maddeleri tamamlanmadan paketi tüm hedef ortamlarda doğrulanmış olarak tanımlamayın.
