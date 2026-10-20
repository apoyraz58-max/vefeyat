# Windows masaüstü Word kurulumu

Vefeyât bir Office görev bölmesi eklentisidir. İlk hedef Windows masaüstü Word’dür. Belgede ad seçme, doğru şahsı arama ve kaynak verideki tarih ifadesini seçili adın sonuna ekleme akışını sağlar. Uygulama bir bulut geliştirme ortamında hazırlanmıştır; gerçek Windows Word’de aşağıdaki kabul kontrolleri ayrıca yapılmalıdır. Tarayıcı ve API kontrollerinin geçmesi Word uyumluluğunu tek başına kanıtlamaz.

## Sunucuyu hazırlama

1. Django uygulamasını ve statik dosyaları Windows bilgisayarından erişilen **HTTPS** adresinde yayımlayın. Örneğin `https://vefeyat.example.org`. Alan adı, HTTPS ve dağıtım yapılandırması ayrı olarak sağlanmalıdır; bulut geliştirme sunucusu kendiliğinden herkese açık bir Word eklentisi sunucusu olmaz.
2. HTTPS sertifikası Windows ve Office tarafından güvenilir olmalıdır. Üretimde güvenilen bir sertifika otoritesinin sertifikasını kullanın. Yerel geliştirmede yalnızca sahibi olduğunuz geliştirme sertifikasını uygun Windows güven deposuna kurun; sertifika ve TLS doğrulamasını kapatmayın. `localhost` Windows bilgisayarını ifade eder, uzaktaki bulut ortamını göstermez.
3. Django’nun host ve ters vekil ayarlarını gerçek alan adına göre yapılandırın. TLS bir ters vekilde sonlanıyorsa gelen şema bilgisini yalnızca güvenilen vekilden kabul edin. Manifestin kaynak adreslerinde `https://` ve doğru alan adı görünmelidir. `WORD_BASE_URL` ortam değişkenini bu kök adresle eşleştirin; varsayılan `https://localhost:8443` yalnızca yerel geliştirme içindir.
4. Windows tarayıcısında `/word/`, `/api/people/?q=ibn&page=1&page_size=20` ve `/static/vefeyat/icon-32.png` adreslerini açın. Manifesti **`/word/manifest.xml`** adresinden indirin; dosya adını örneğin `vefeyat.xml` yapın. XML içindeki adreslerin sunucunun gerçek HTTPS adresi olduğunu kontrol edin.
5. Office.js için `https://appsforoffice.microsoft.com` erişilebilir olmalıdır. API ve görev bölmesi aynı kökte çalışır; eklentiye ayrıca API anahtarı girilmez. Word içindeki gömülü tarayıcı ağ ve sertifika erişimine ayrıca ihtiyaç duyar.

Yerel geliştirmede HTTP üzerinden `/word/` sayfasında manuel arama yapılabilir. Word’e yüklemek için güvenilir HTTPS bağlantısı hazırlayın. Manifestteki `ReadWriteDocument` izni yalnızca kullanıcının başlattığı seçim okuma ve tarih ekleme işlemleri için kullanılır; belgenin tamamı API’ye gönderilmez. Arama servisine arama alanındaki ad gönderilir.

## Windows’ta deneme amaçlı yükleme

Microsoft 365’in güncel masaüstü Word sürümü önerilir. En az **WordApi 1.3** gerekir. Bu sürüm, aynı adın belgenin başka bir yerinde seçilmesini ayırt etmek için gereken konum karşılaştırmasını sunar. Daha düşük sürümlerde belgeye ekleme kapalıdır; manuel arama kullanılabilir. Office sürümü, kurum ilkeleri ve güncelleme kanalı nedeniyle menü adları değişebilir. Kurumsal yüklemelerde yöneticinin eklenti dağıtım yöntemi kullanılmalıdır; aşağıdaki paylaşılan katalog yöntemi geliştirme ve deneme içindir.

1. İndirdiğiniz `vefeyat.xml` dosyasını Windows bilgisayarında size ait bir klasöre koyun. Dosya Gezgini’nde klasör özelliklerinden paylaşılan bir klasör oluşturun. Bu işlem için yerel Windows izinleri gerekebilir; parolaları uygulamaya veya sohbet ekranına yazmayın.
2. Klasörün UNC adresini not edin; örneğin `\\BILGISAYARADI\VefeyatEklentileri`. Bu klasör yalnızca manifesti barındırır; web uygulaması HTTPS sunucusunda kalır. Manifest klasörünü internete açmanız gerekmez.
3. Word’de **Dosya → Seçenekler → Güven Merkezi → Güven Merkezi Ayarları → Güvenilir Eklenti Katalogları** yolunu açın. Katalog adresine UNC klasörünü ekleyin; **Menüde göster** seçeneğini işaretleyin.
4. Word’ü tamamen kapatıp yeniden açın. **Ekle → Eklentilerim → Paylaşılan Klasör** bölümünden Vefeyât’ı ekleyin. Güncel sürümlerde bu seçenek **Giriş → Eklentiler → Diğer Eklentiler** altında olabilir.
5. Eklenti izinlerini inceleyin. **Giriş** sekmesindeki **Vefeyât → Şahıs ara** düğmesiyle görev bölmesini açın. Katalog bölümü görünmüyorsa kurumun Office eklenti ilkelerini ve Word sürümünü kontrol edin; güvenlik kısıtlamalarını devre dışı bırakmayın.

Microsoft’un ilgili yönergeleri:

- [Windows’ta ağ paylaşımından Office eklentisi yükleme](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Office eklentilerinin barındırılması ve dağıtımı](https://learn.microsoft.com/office/dev/add-ins/publish/publish)
- [Word API gereksinim kümeleri](https://learn.microsoft.com/javascript/api/requirement-sets/word/word-api-requirement-sets)
- [Word.Range: konum karşılaştırması ve aralık takibi](https://learn.microsoft.com/javascript/api/word/word.range?view=word-js-1.3)
- [Microsoft OfficeDev: Word.Range referans kaynağı](https://github.com/OfficeDev/office-js-docs-reference/blob/live/docs/docs-ref-autogen/word/word/word.range.yml)

## Kullanım

1. Belgeye `İbn Sînâ` yazın; yalnızca adı seçin.
2. **Word’deki seçili adı al** düğmesine basın. Seçim okunur ve arama başlar.
3. Tarih, tanıtım ve kaynak bağlantısından doğru şahsı seçin. Aynı adı taşıyan kişiler ayrı sonuçlardır.
4. **Tarih ifadesini ekle** düğmesine basın. Ad değiştirilmez; sonuna bir boşluk ve kayıttaki `date_text` aynen eklenir. Örneğin `İbn Sînâ (ö. 428/1037)`.
5. Ekleme sonrası düğme yeni bir Word seçimi alınana kadar kapanır. Yanlışlıkla eklenen metni Word’ün **Geri Al** komutuyla kaldırın.

Elle arama yapılabilir; arama kutusunu düzenlemek daha önce alınan Word seçimini geçersiz kılar. Belgeye eklemek için yeniden Word’den ad seçilmelidir. Word dışında `/word/` sayfasında arama ve kaynak açma çalışır; belgeye ekleme kapalıdır.

Tarih alanındaki yaklaşık tarihler, yüzyıllar ve `gaybeti` gibi ifadeler özgün biçimiyle eklenir; bütün kayıtlar vefat tarihiymiş gibi yeniden yorumlanmaz. Kullanıcı kaydı ve bağlamı kontrol eder. Eklenti seçilen adın devamındaki Türkçe eki veya noktalama işaretini yeniden düzenlemez.

## Kabul kontrolleri

Deneme belgesinde şu durumları Windows Word içinde kontrol edin:

- `İbni Sina` gibi bir ad varyantından doğru kaydı bulma; aynı adı taşıyan iki kişiyi ayırt etme.
- Kaynak bağlantısının açılması ve seçilen kaydın özgün tarih ifadesinin eklenmesi.
- Boş seçimde hiçbir metin yazılmaması.
- Aramadan sonra farklı bir ad seçildiğinde eklemenin durması ve yeniden seçim istenmesi.
- Aramadan sonra belgenin başka yerindeki **aynı ad** seçildiğinde de eklemenin durması.
- Manuel arama sonrasında yeni Word seçimi alınmadan ekleme yapılamaması.
- Tek eklemenin Word’ün Geri Al komutuyla geri alınması; belge biçiminin kontrolü.
- Tarih zaten seçili metinde veya tek eşleşmeli paragrafta adın hemen sonrasındaysa tekrar eklenmemesi.
- API erişimi kesildiğinde anlaşılır hata ve tekrar arama; görev bölmesini yeniden açma.
- Yaklaşık tarih, yalnızca yüzyıl ve gaybet gibi kesin vefat tarihi olmayan ifadelerin değişmeden korunması.

Seçim alındığında Word aralığı `context.trackedObjects.add(range)` ile izlenir. Ekleme anında aynı bağlam `Word.run(savedRange, ...)` ile kullanılır; güncel seçimin metni özgün metinle, konumu ise `compareLocationWith(savedRange)` sonucundaki `Word.LocationRelation.equal` ile karşılaştırılır. Metin veya konum değiştiyse hiçbir şey yazılmaz. Konum karşılaştırması ya da aralık takibi başarısız olursa metne göre ekleme yapan bir yedek akışa geçilmez. Seçim geçersiz kılındığında veya işlem bittiğinde izlenen aralık serbest bırakılır.

Mevcut tarih soneki denetimi, seçimin kendi metninde veya aynı adın paragrafta tek geçtiği durumda uygulanır. Aynı paragrafta birden fazla aynı ad varsa bu ek kontrol atlanır; yinelenen eklemeleri azaltan kontroller tüm belge için tekrar temizleme sağlamaz. Tarih zaten eklendiyse yeniden seçip eklemeyin.

## Tarayıcı otomasyonu ve Office API taklidiyle ek kontroller

`node --test scripts/tests/word.test.mjs` komutu Office/API taklit testlerini çalıştırır. `/word/` sayfası Office.js olmadan da test edilebilir. Tarayıcı otomasyonunda yalnızca Office.js isteğini test taklidiyle değiştirip `Office.onReady`, `Office.HostType.Word`, `Office.context.requirements.isSetSupported`, `Word.LocationRelation.equal` ve `Word.run` nesnelerini sağlayın. `Word.run` bağlamında `document.getSelection()`, `selection.load()`, `selection.compareLocationWith()`, `selection.paragraphs.load()`, `context.trackedObjects.add/remove`, `context.sync()`, `selection.insertText()` ve dönen aralığın `select()` işlevlerini taklit edin. `Word.run(savedRange, callback)` overloadu, saklanan aralığın bağlamını kullanmalıdır. Ağ yanıtı gerçek API ile aynı `{query,count,page,page_size,results,has_next}` biçimini taşımalıdır.

Özellikle değişen seçimde `insertText` çağrılmadığını, başarılı eklemede yalnızca `" " + date_text` metninin `End` konumuna yazıldığını, ikinci tıklamada yeni seçim olmadan yazılmadığını, kötücül HTML içeren API alanlarının düz metin gösterildiğini ve `javascript:` kaynak bağlantılarının üretilmediğini doğrulayın. Bunlar gerçek Office köprüsünü, sertifika güvenini ve Word’ün Geri Al davranışını doğrulamaz; Windows kabul kontrolleri yine gereklidir.
