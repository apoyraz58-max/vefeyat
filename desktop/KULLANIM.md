# Vefeyât — Windows Word'de çevrimdışı kullanım

Vefeyât, Windows masaüstü Word'e **Vefeyât** sekmesi ekler. Şahısları bilgisayarınızdaki kayıtlarda arar; seçtiğiniz kaydın vefat tarihini Word'de seçili adın sonuna ekler. Arama, düzenleme, yedekleme ve belgeye ekleme için internet, sunucu, alan adı veya Docker gerekmez.

## Kurulum

Windows 10 veya 11, Windows masaüstü Microsoft Word ve bilgisayarda bulunan .NET Framework 4.x gerekir; .NET Framework 4.8 önerilir. Paket 32 bit ve 64 bit Word için hazırlanmıştır. Word'ün web ve Mac sürümleri bu eklentiyi çalıştırmaz.

1. ZIP paketindeki **tüm dosyaları**, örneğin `C:\Vefeyat-Kurulum` gibi yerel ve OneDrive ile eşitlenmeyen bir klasöre çıkarın. Dosyaları ZIP içinden çalıştırmayın.
2. Bütün Word pencerelerini kapatın.
3. Çıkardığınız klasörde **Kur.cmd** dosyasına çift tıklayın. Normal kullanıcı hesabınızla çalıştırın; yönetici olarak açmanız gerekmez.
4. Kurulumun tamamlandığını bildiren mesajdan sonra Word'ü açın. Şeritte **Vefeyât** sekmesini göreceksiniz.

Kurulum bilgisayardaki Windows bileşenlerini kullanır; internetten dosya veya ek geliştirme aracı indirmez. İlk katalog açılışında paketle gelen kayıtlar kişisel veri klasörünüze alınır. Size hazırlanan veri paketinde 7.934 kayıt bulunur; kaynak kodundaki örnek CSV ise üç hayalî kayıt içerir. Kurulumu tekrar çalıştırmak mevcut şahıs kayıtlarınızı sıfırlamaz.

## Belgeye vefat tarihi ekleme

1. Düzenlenebilir Word belgesinde yalnız şahıs adını seçin; paragraf sonunu veya tablo hücresi sonunu seçime katmayın.
2. **Vefeyât → Seçili adı ara** düğmesine basın.
3. Sonuçlardan doğru şahsı seçin. Tam adı, orijinal adı, tanıtımı ve **Eklenecek vefat tarihi** alanını kontrol edin.
4. **Vefat tarihini belgeye ekle** düğmesine basın.

Örneğin seçili `Abaka` adı, ilgili kayıt seçildiğinde `Abaka (ö. 680/1282)` olur. Şahıs adı ve mevcut biçimlendirmesi korunur. Tarih ifadesi takvim dönüşümü veya tarih tahmini yapılmadan kayıttaki metin olarak eklenir; ifadenin çevresindeki boşluklar temizlenir.

Word'de **Ctrl+Z** ile eklemeyi geri alabilirsiniz. Aynı tarih adın hemen sonrasında zaten varsa eklenti tekrar eklemez. Belge veya seçim değişmişse işlem durur; adı yeniden seçerek başlayın. Boş veya çok satırlı tarihler ve düzenlemeye kapalı belgeler için ekleme yapılmaz.

## Şahıs kayıtlarını kullanma

**Vefeyât → Şahısları ara**, belgeye bir şey eklemeden yerel kataloğu açar. Adın bir bölümünü, tanıtımda geçen bir kelimeyi veya orijinal adı arayabilirsiniz. `maturidi` gibi sade yazımlar, `Mâtürîdî` gibi işaretli yazımları da bulur. Boş arama tüm kayıtları listeler; ilk 200 sonuç gösterilir. Daha az sonuç için aramayı daraltın.

- **Yeni**, yeni şahıs kaydı açar. Kısa ad zorunludur.
- **Düzenle**, seçili kaydı değiştirir; **Kaydet** ile bilgisayara kaydeder.
- **Sil**, onayınızdan sonra yalnız seçili kaydı siler.

Tarih alanına belgeye eklenmesini istediğiniz ifadeyi, örneğin `(ö. 333/944)`, yazın. Kaynak adresleri yalnız metin olarak saklanır ve görüntülenir; kaynak siteler açılmaz veya internetten içerik indirilmez. Arapça ve diğer Unicode metinler kayıtlarda korunur.

İki Word işlemi aynı veriler üzerinde çalışıyorsa eski veriye dayanan bir değişiklik reddedilebilir. Böyle bir mesaj aldığınızda katalog penceresini kapatıp yeniden açın, güncel kaydı kontrol edin ve değişikliğinizi yeniden uygulayın.

## CSV içe ve dışa aktarma

**CSV içe aktar** için dosya UTF-8 kodlamasında olmalı ve ilk satırda aşağıdaki altı başlık aynı sırada bulunmalıdır:

```text
Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link
```

Noktalı virgül ayracı kullanın; virgülle ayrılmış dosyalar da kabul edilir. İçinde ayraç, tırnak veya satır sonu bulunan alanlar CSV kurallarına göre çift tırnak içinde olmalıdır. Excel'de kaydederken **CSV UTF-8** biçimini seçin. `---` değerleri boş alan sayılır. Dosyada geçersiz bir kayıt varsa aktarımın tamamı durdurulur; önceki satırlar kısmen kaydedilmez.

Aynı dosyayı yeniden aktarmak aynı kayıtları çoğaltmaz. Kısa ad, tam ad, orijinal ad ve kaynak adresi aynı olan tek bir kayıt varsa tarih veya tanıtım değişikliği o kaydı günceller. Birden çok kayıt aynı kimlik bilgilerini taşıyorsa eklenti hangi şahsın değiştiğini tahmin etmez; eşleştirilemeyen içerikleri ayrı kayıt olarak korur. Ad veya kaynak adresi değişmiş bir satır da yeni kayıt sayılabilir. Kimlik bilgilerini değiştirmek için **Düzenle** kullanın ve büyük bir aktarımdan önce yedek alın.

**CSV dışa aktar**, arama sonucundan bağımsız olarak bütün kayıtları altı sütunlu, Excel ile uyumlu UTF-8 CSV dosyasına çıkarır. Kayıtların uygulama içindeki kimlikleri CSV'ye yazılmaz; tam yedek için XML yedeğini kullanın.

## Verilerin yeri ve yedekleme

Şahıs kayıtlarınız şu yerel dosyada saklanır:

```text
%LOCALAPPDATA%\Vefeyat\Data\catalog.xml
```

**Vefeyât → Veri klasörü** bu klasörü açar. Eklentinin kendi dosyaları ayrı olarak `%LOCALAPPDATA%\Vefeyat\Addin` altında bulunur. Paketteki başlangıç CSV'sini değiştirmek mevcut kişisel kayıtlarınızı güncellemez.

**Yedek al**, bütün kayıtları ve kimliklerini bir XML dosyasına kaydeder. Yedek için bilgisayarınızda veya harici diskinizde yerel bir konum seçin. **Yedeği geri yükle**, seçtiğiniz geçerli XML yedeğini onayınızdan sonra mevcut kayıtların yerine koyar. Geri yüklemeden önce güncel verilerin ayrı bir yedeğini alın. Bu işlem Word belgelerinin içeriğini değiştirmez.

Her başarılı veri değişikliğinde önceki sürüm `catalog.previous.xml` olarak aynı klasörde korunur. Bu dosya her değişiklikte yenilendiği için düzenli olarak ayrı XML yedekleri alın. Yerel veri dosyası okunamadığı için katalog açılamıyorsa eklenti bir XML yedeğinden kurtarma seçeneği sunar. Önce `Data` klasörünün tamamını başka bir yerel klasöre kopyalayın; sonra kurtarma sorusunu onaylayıp geçerli XML yedeğinizi seçin. Önceki sürüm dosyasını kullanacaksanız onun da ayrı bir kopyasını alın: kurtarma işlemi mevcut çalışma dosyasını `catalog.previous.xml` olarak saklar ve eski önceki-sürüm dosyasını yeniler. İptal ederseniz veri dosyanız değiştirilmez.

## Güncelleme ve kaldırma

Onarım veya yeni paketi kurmak için Word'ü kapatıp **Kur.cmd** dosyasını yeniden çalıştırın. Kişisel veri klasörü korunur.

Kaldırmak için bütün Word pencerelerini kapatıp **Kaldir.cmd** dosyasını çalıştırın. Eklenti Word'den kaldırılır; `%LOCALAPPDATA%\Vefeyat\Data` altındaki şahıs kayıtları ve yedekleriniz korunur. Yeniden kurduğunuzda mevcut kayıtlar kullanılır.

## Sekme görünmüyorsa

Word'ü tamamen kapatıp yeniden açın. **Dosya → Seçenekler → Eklentiler** altında **Yönet: COM Eklentileri → Git** yolundan **Vefeyât** kaydını kontrol edin. Word eklentiyi devre dışı bıraktığını bildiriyorsa aynı sayfadaki **Devre Dışı Bırakılan Öğeler** listesini inceleyin ve bildirilen nedeni giderin.

Bu paket imzasız bir COM eklentisidir. Kurumunuzun imza veya eklenti politikası kurulumu ya da yüklemeyi engelliyorsa kurumunuzun BT biriminden onaylı kurulum veya imzalı dağıtım isteyin. Güvenlik korumalarını kapatmayın.

Önceki web tabanlı Vefeyât eklentisi hâlâ kuruluysa iki Vefeyât sekmesi görebilirsiniz. Eski eklentiyi Word'ün **Ekle → Eklentilerim** bölümünde ilgili eklentinin menüsünden kaldırın; menü adı Word sürümüne göre **Eklentiler** veya **Eklentilerim** olabilir. Eski eklenti kurumsal olarak dağıtıldıysa kaldırma işlemi için BT biriminizi kullanın. Yeni yerel eklenti eski web eklentisini otomatik olarak kaldırmaz.
