# Vefeyât kullanım kılavuzu (Word eklentisi, 0.4)

Vefeyât, metinde geçen şahıs adlarını DİA kataloğunda bulur, ilk geçişlerine ölüm
tarihini ekler, yazımlarını denetler ve teslimden önce belgeyi toparlar. Bu sürüm
VBA eklentisinden bağımsızdır; VBA sekmesindeki bütün komutları içerir.

Yeni bir özelliği önce bir tez **kopyasında** deneyin.

## 1. Kurulum ve güncelleme

Manifest: `https://apoyraz58-max.github.io/vefeyat/manifest.xml`

**Windows masaüstü Word:** Manifesti bir klasöre koyun ve klasörü paylaşın (sağ tık →
Özellikler → Paylaşım). Word → **Dosya → Seçenekler → Güven Merkezi → Güven Merkezi
Ayarları → Güvenilen Eklenti Katalogları** → paylaşılan klasörün `\\BILGISAYAR\klasor`
yolunu ekleyin, "Menüde göster"i işaretleyin, Word'ü yeniden başlatın → **Ekle →
Eklentilerim → Paylaşılan Klasör → Vefeyât**.

**Mac Word:** Manifesti `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
klasörüne kopyalayın (yoksa oluşturun), Word'ü yeniden başlatın → **Ekle → Eklentilerim**.

**Word Online:** **Giriş → Eklentiler → Eklentilerim → Eklentimi Karşıya Yükle** →
`manifest.xml`.

**Güncelleme.** Panelin içindeki değişiklikler kendiliğinden gelir: panel her açılışta
yayındaki sürümü kontrol eder, elindeki kopya eskiyse kendini yeniden yükler. Panelin
altında sürüm numarası yazar (örnek: 0.4.0.47). Şeride ya da sağ tık menüsüne yeni
düğme eklendiğinde manifesti bir kez yenilemek gerekir: yeni `manifest.xml` dosyasını
indirip paylaşılan klasördekinin yerine koyun, Word'ü kapatın,
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` klasörünü silin, Word'ü açın.

**Klavye kısayolu (deneme).** Ayrı bir manifest "Ölüm tarihi ekle"yi
**Ctrl+Alt+Shift+V** (Mac: Cmd+Option+Shift+V) kısayoluna bağlar:
`https://apoyraz58-max.github.io/vefeyat/manifest-kisayol.xml`. Word kısayola yalnız
"paylaşılan çalışma ortamı"nda izin verir; bu yüzden bu sürümde bütün şerit düğmeleri
paneli yeniden yüklemek yerine açık panele komut gönderir. Ayrı bir eklenti olarak
görünür ("Vefeyât (kısayol denemesi)", sekmesi "Vefeyât (deneme)"); asıl eklentiye
dokunmaz. Windows Word 2408 ve sonrası gerekir. Denerken asıl `manifest.xml`
dosyasını paylaşılan klasörden geçici olarak çıkarın, yoksa iki Vefeyât sekmesi görünür.
Kısayol başka bir komutla çakışırsa Word bir kez hangisinin çalışacağını sorar.

## 2. Hızlı başlangıç

1. Metinde bir şahıs adını seçin (çift tıklama yeter; Türkçe ek sorun değildir).
2. **Vefeyât → Ölüm tarihi ekle**. Ad tek ve kesin eşleşiyorsa, ilk geçişse ve
   arkasında tarih yoksa tarih hemen eklenir. Aksi hâlde sağdaki pencere ne
   yapılacağını sorar.
3. Bütün belge için **Şahıs dökümü** ya da **Tüm belgeyi denetle**.

## 3. Şerit

| Grup | Komutlar |
| --- | --- |
| Ekle ve denetle | Ölüm tarihi ekle, Tüm belgeyi denetle, İsim yazımı |
| Denetim | Tarih ekleme ▾, Denetlenecek ▾, Denetime devam et, Toplu düzeltme önizlemesi |
| Şahıs ve kaynak | Şahıs ara, DİA maddesini aç, Tarihin kaynağı |
| Belge | Şahıs dökümü, Diğer araçlar ▾, Belge sağlığı, İşlem geçmişi |
| Teslim ve dizin | Tutarlılık raporu, Şahıs raporu, Şahıslar dizini ▾, Teslim öncesi temizle |
| Tercihler | Ad biçimi ▾, Kısaltma ▾, Ayarlar ▾ |
| Veri ve yardım | Veri bilgisi, CSV seç, Kullanım kılavuzu |

Şeritteki menüler seçili değeri gösteremez (Office eklentilerine izin verilmez);
geçerli ayarlar panelin en üstünde yazar.

**Sağ tık:** Seçili adda **Vefeyât → Ölüm tarihi ekle / İsim yazımı / DİA maddesini
aç / Tarihin kaynağı**. Kırmızı çizgili (yazım hatası sanılan) kelimede Word kendi
yazım menüsünü açar ve eklentiler oraya komut ekleyemez. **Tercihler → Ayarlar →
Şahıs adlarını yazım denetiminden çıkar** ile çizgi kalkar; tarih eklediğiniz adlar
kendiliğinden çıkarılır.

## 4. Tarih ekleme

**Tarih ekleme menüsü:**
- *Yalnız ilk kullanım* (varsayılan): kişi metinde daha önce geçtiyse uyarı çıkar
  (**Önceki geçişe git** / **Yine de ekle**). Uyarı yazıma değil kimliğe bakar:
  önceki "Fahreddin er-Râzî", Ebû Bekir er-Râzî için önceki geçiş sayılmaz; önce
  "Gazali" diye yazılmış geçiş Gazzâlî için sayılır.
- *Her bölümde ilk kullanım*: aynı kural her bölümde baştan işler; kişinin her
  bölümdeki ilk geçişi tarihi alır. Bölüm başı, Word'ün **Başlık 1** biçemi ya da
  "BİRİNCİ BÖLÜM", "2. Bölüm", "GİRİŞ", "SONUÇ" gibi bir başlıkla tanınır.
- *Bütün kullanımlar*: uyarı çıkmaz.
- *Yalnız denetle*: tarih eklenmez.

Ana metin ile dipnot/sonnotlar **ayrı** değerlendirilir: kişinin ana metindeki ilk
geçişi de, notlardaki ilk geçişi de tarihi alır; "daha önce geçti" uyarısı yalnız aynı
bölümdeki (metin ya da notlar) önceki geçişlere bakar.

**İlk geçiş bakımı.** Tarihi ilk geçişe ekledikten sonra metnin önüne aynı kişiyi
yazarsanız, ya da bir geçişte gereksiz bir tarih kalırsa, Vefeyât bunu fark eder ve
pencerede "İlk geçiş değişti" bildirimi gösterir. **Gözden geçir** Toplu düzeltme
önizlemesini açar: tarih yeni ilk geçişe *taşınır* (yazdığınız biçimiyle), sonraki
geçişteki *kaldırılır*; ilk geçişte zaten tarih varken sonrakilerdeki fazla tarihler
de kaldırılmak üzere önerilir. Elle yazılmış tarihler de tanınır; yalnız katalogla
uyuşanlar taşınır ya da kaldırılır, kimliği tahmine dayananlar işaretsiz gelir.
**Onayınız olmadan hiçbir şey değişmez.** (Bildirim, belgeyi değişiklik olaylarıyla
izleyebilen güncel Microsoft 365 Word'de kendiliğinden çıkar; diğerlerinde Toplu
düzeltme önizlemesini açtığınızda öneriler yine görünür.)

Adın arkasında zaten tarih varsa ikinci tarih eklenmez. **Kısaltma** menüsü önek
biçimini (ö. / öl. / v. / kaynaktaki gibi) belirler.

**Eş adlılar.** Bir ad birden çok kişiye uyuyorsa pencere adayları gösterir; seçili
olan renkli ve "seçili" etiketlidir. Seçiminiz o geçiş için belgeye kaydedilir
(Diğer araçlar → Kimlik eşleşmelerini düzelt). Vefeyât önce şunlara bakar ve
satırda nereden bildiğini yazar:
- **yazılı tarih:** "Râzî (ö. 313/925)" → Ebû Bekir er-Râzî;
- **belgedeki diğer geçişler:** belgede yalnız bir Gazzâlî tam adıyla geçiyorsa kısa
  "Gazzâlî"ler ona bağlanır;
- **cümlenin konusu:** belgede iki Râzî de geçiyorsa "Râzî tefsirinde…" → Fahreddin,
  "Râzî tıpta…" → Ebû Bekir (cümlede konu yoksa paragrafın konusu);
- **baş harf:** "F. er-Râzî" → Fahreddin er-Râzî.

"Doğrulayın" yazan satırlar tahmindir; tarih eklediğinizde onaylanmış olur.
Benzer yazılışlılar arasında tek bir tam eşleşme varsa ("İbn Sînâ" ile
"İbn Sinân Hafâcî") tam eşleşen kendiliğinden seçilir. Arapça takı (el-/er-/eş-)
karşılaştırmada yok sayılır: "Ebû Bekir er-Râzî" ile "Ebû Bekir Râzî" aynıdır.

**İsim yazımı.** Liste metindeki yazımla başlar ve o seçilidir; başka bir yazım
seçince altında ne yazılacağı görünür ("Gazali'nin → Gazzâlî'nin"). Türkçe ek korunur.
**Ayarlar → Seçili yazımı tercih et** bir kişinin yazımını bu belge için sabitler.

## 5. Bütün belge

- **Şahıs dökümü:** metin, dipnot ve sonnotlardaki bütün geçişler; her satırda
  katalog kaydı, yazılı tarih ve durum. Özet satırı kaç dipnot ve sonnot
  okunduğunu da yazar. Satıra tıklayınca geçiş belgede seçilir.
  **Excel'e aktar** dökümü .xlsx dosyası olarak kaydeder (başlık satırı sabit,
  süzgeçli); **Panoya kopyala** Excel'e ya da bir Word tablosuna yapıştırılır.
- **Tüm belgeyi denetle:** geçişleri tek tek gezer; **Denetimi bitir** duraklatır,
  **Denetime devam et** kaldığı yerden sürdürür. *Yalnız ilk kullanım* ayarında
  tekrarlar atlanır; yanlış yazılmış tarih olan geçişte her zaman durur.
- **Toplu düzeltme önizlemesi:** eksik ilk geçiş tarihleri, katalogdan farklı
  yazımlar ve katalogla çelişen tarihler tek listede. Yalnız işaretliler uygulanır;
  tarih düzeltmeleri işaretsiz gelir.
- **İşlem geçmişi:** bu oturumdaki ekleme, düzeltme ve onarımlar.

Büyük belgede belge bir kez okunur; sonra yalnız değişen paragraflar yeniden okunur
(güncel Microsoft 365). Uzun işlemlerde çalışma çubuğu, geçen süre ve **İptal** görünür.

## 6. Neler şahıs sayılmaz

- Mezhep, tarikat, hanedan ve bölge sıfatları: "Nakşibendî tarikatı", "Sünnî ve Şiî
  müellifler", "Abbâsî halifesi", "Hanefî fıkhı", "Osmanlı âlimleri Mâtürîdî idi".
  Tek başına Şâfiî, Eş‘arî ve Mâtürîdî imamın kendisidir.
- Sıfat olarak kullanılan adlar: "Usûlî açıdan", "Lügavî anlamı"; belgede küçük
  harfle de geçen kelime cümle başında ya da başlıkta.
- Uzun adlardan kopan yer, unvan ve kelime parçaları: "Irak", "Arap", "Has", "Cem",
  "Hocası" (tam ad "Cem Sultan" yine bulunur).
- Atıflar: "bk. Gazzâlî, el-Mustasfâ, I, 45", "(Gazzâlî, 1993: 45)", "Gazzâlî, age,
  II, 45" — metinde de dipnotta da; dipnotun anlatım kısmı okunur.
- Kaynakça: "Kaynakça/Bibliyografya/Kaynaklar" başlığından "Ekler/Dizin/Özgeçmiş"
  gibi bir başlığa kadar, ve kaynakça girdisi biçimindeki paragraflar.

Yine de yanlış sayılan bir ifade için seçili ad penceresinde **Şahıs değil olarak
işaretle**; o belgede bir daha listelenmez (Diğer araçlar → Belge istisnaları).

## 7. Tarih denetimi

- **Elle yazılmış tarihler:** adın arkasındaki tarih katalogdakinden farklıysa ya da
  hicrî ve milâdî yıl birbirini tutmuyorsa (505/1111 doğru, 505/1211 yanlış) dökümde,
  denetimde ve şahıs raporunda uyarı çıkar. Önek farkı, yalnız bir takvim ya da
  kısaltılmış yıl ("917-18") uyarı vermez.
- **Belge sağlığı:** Vefeyât'ın eklediği tarihleri katalog ve önek ayarıyla
  karşılaştırır; seçtiklerinizi onarır. Önek değiştirdiyseniz eski önekli tarihler
  burada hazır işaretli gelir.
- **Tarihin kaynağı:** imleçteki Vefeyât tarihinin hangi kayıttan geldiğini gösterir.
- **Tutarlılık raporu:** ö./öl./v. ve tarih biçimlerinin karışık kullanıldığı yerler.

## 8. Teslim

- **Şahıs raporu:** şahıslar ilk geçiş sırasıyla, notlarla ve milâdî yıla göre
  kronoloji; yeni bir Word belgesi olarak açılır.
- **Şahıslar dizini:**
  - Word sayfa bilgisini okuyabiliyorsa (güncel Microsoft 365 masaüstü) dizin,
    belgeye işaret konmadan sayfa numaralarından hızlıca hazırlanır. Panelde
    "1 numaralı sayfa"yı ayarlayın ya da imleci o sayfaya koyup **İmlecin olduğu
    sayfa**ya basın; öncesi Roma rakamıyla yazılır. Numaralar sabittir; metin
    değişince yeniden oluşturun.
  - Aksi hâlde gizli dizin işaretleriyle hazırlanır; büyük belgede uzun sürer.
    "Dipnotları da kat" kutusunu kapatmak süreyi kısaltır.
  - **Dizin işaretlerini kaldır:** dizin güncel hâliyle metin olarak kalır.
    **Dizini tamamen kaldır:** dizin, başlığı ve işaretler birlikte gider.
- **Teslim öncesi temizle** (kopyada!): tarihler ve dizin metin olarak kalır;
  Vefeyât denetimleri, gizli dizin işaretleri, onaylar, istisnalar, tercihler ve
  kendi şahıs kayıtlarınız belgeden kaldırılır.

## 9. Kendi şahıslarım ve tercihler

**Kendi şahıslarım:** katalogda olmayan kişiler (modern yazarlar, DİA'da maddesi
olmayanlar). **Veri bilgisi** penceresindeki bölümden ya da seçili ad katalogda yoksa
çıkan **Kendi şahıslarıma ekle** düğmesiyle ekleyin: kısa ad ve ölüm tarihi ("1998",
"ö. 505/1111", "1913-1998"), isteğe bağlı tam ad ve açıklama. Katalog kaydı gibi
kullanılır ("kendi kaydınız" notuyla).

**Tercihleri dışa aktar / al:** ayarlar, yazım tercihleri, istisnalar, kimlik
onayları ve kendi şahıslarınız bir dosyada; başka bir tezde aynen kullanılabilir.

**Katalog:** eklentiyle gelen DİA kataloğu kendiliğinden yüklenir; **CSV seç** ile
kendi dosyanızı kullanabilirsiniz.

## 10. VBA sürümünden geçiş

VBA sürümünün eklediği tarihler Word alanıdır. **Belge sağlığı** bunları sayar ve
**VBA tarihlerini dönüştür** ile aynı metni taşıyan Vefeyât tarihlerine çevirir. Kişi
tarihin önündeki addan, eş adlılarda tarihin kendisinden bulunur; önündeki ad
bulunamayan alana dokunulmaz. Dönüştürdükten sonra VBA eklentisini kaldırabilirsiniz.

## 11. Bilinen sınırlar

- Üst/alt bilgiler, açıklamalar ve metin kutuları taranmaz (Diğer araçlar → Denetim
  kapsamını göster).
- Küçük harfle yazılmış adlar ("gazali") bulunmaz.
- Panel açıkken şerit düğmesine yeniden basmak paneli yenilemeyebilir; o zaman
  penceredeki düğmeyi kullanın.
- Şahıs raporu yeni belge açar; Word Online'da açılır pencere engelleyicisi buna
  izin vermeyebilir.

## 12. Deneme tezi ve geri bildirim

Bütün komutları denemek için `https://apoyraz58-max.github.io/vefeyat/vefeyat-deneme.docx`
dosyasını açın; ilk sayfadaki kontrol listesini izleyin.

Sorun görürseniz: panelin altındaki sürüm numarası, ne yaptığınız ve paneldeki
mesajın ekran görüntüsü.

## Hakkında

**Vefeyât** — Geliştiren: **Abdurrahman Poyraz**, Sivas Cumhuriyet Üniversitesi
İlahiyat Fakültesi. Şahıs bilgileri ve ölüm tarihleri TDV İslâm Ansiklopedisi (DİA)
kataloğundan alınır. Simgeler: Fluent UI System Icons (MIT).
