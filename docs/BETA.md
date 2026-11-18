# Vefeyât kapalı beta: kurulum ve deneme

Vefeyât, Word'de şahıs adlarını TDV İslâm Ansiklopedisi (DİA) kataloğuyla karşılaştıran,
ilk geçişe ölüm tarihini ekleyen ve adların yazımını denetleyen bir eklentidir. Bu sayfa
kapalı beta denemecileri içindir.

- **Geliştiren:** Abdurrahman Poyraz, Sivas Cumhuriyet Üniversitesi İlahiyat Fakültesi
- **Sorun ve öneriler:** apoyraz@cumhuriyet.edu.tr
- **Gizlilik:** Belge metniniz bilgisayarınızdan çıkmaz. Eklenti yalnız kendi dosyalarını ve
  kataloğu internetten indirir; tezinizi hiçbir yere göndermez.

> Önce bir **kopya** üzerinde deneyin. Vefeyât'ın yaptığı her değişiklik Ctrl+Z ile geri
> alınabilir, ama beta sürümünde beklenmedik bir durum olabilir.

## 1. Hangi manifest?

Manifest, Word'e eklentiyi tanıtan küçük bir dosyadır. Word sürümünüze göre birini seçin:

| Word | Dosya |
|---|---|
| Microsoft 365, sürüm **2408 veya sonrası** (Windows) / **16.88 veya sonrası** (Mac) | [manifest.xml](https://apoyraz58-max.github.io/vefeyat/beta/manifest.xml) |
| Daha eski Word, Office 2019/2021 | [manifest-uyumlu.xml](https://apoyraz58-max.github.io/vefeyat/beta/manifest-uyumlu.xml) |

Sürümünüzü görmek için: Word → **Dosya → Hesap → Word Hakkında**. En üstte "Sürüm 2410"
gibi bir numara yazar. Emin değilseniz önce `manifest.xml` ile deneyin; Word eklentiyi
kurmazsa `manifest-uyumlu.xml` kullanın.

İkisi arasındaki tek fark klavye kısayollarıdır (Alt+Shift+V ile tarih ekleme gibi). Uyumlu
sürümde kısayol yoktur; bütün komutlar şeritte ve sağ tık menüsünde aynen bulunur.

Bağlantıya tıklayınca dosya tarayıcıda açılırsa sayfaya sağ tıklayıp **Farklı kaydet** deyin.
Dosyanın adı `manifest.xml` ya da `manifest-uyumlu.xml` olarak kalmalı.

## 2. Windows'ta kurulum

1. Bilgisayarınızda bir klasör açın, ör. `C:\Vefeyat`, ve manifest dosyasını içine koyun.
2. Klasöre sağ tıklayın → **Özellikler → Paylaşım → Paylaş…** → kendinizi seçip
   **Paylaş**'a basın. Pencerede klasörün ağ yolu görünür: `\\BILGISAYARADI\Vefeyat`.
   Bu yolu kopyalayın.
3. Word → **Dosya → Seçenekler → Güven Merkezi → Güven Merkezi Ayarları → Güvenilen
   Eklenti Katalogları**. "Katalog URL'si" kutusuna kopyaladığınız yolu yapıştırın,
   **Katalog Ekle**'ye basın, **Menüde göster** kutusunu işaretleyin, **Tamam** deyin.
4. Word'ü tamamen kapatıp yeniden açın.
5. **Ekle → Eklentilerim** (bazı sürümlerde **Giriş → Eklentiler → Diğer Eklentiler**) →
   **Paylaşılan Klasör** sekmesi → **Vefeyât (beta)** → **Ekle**.

Şeritte **Vefeyât (beta)** sekmesi görünür.

## 3. Mac'te kurulum

1. Finder'da **Git → Klasöre Git…** (Shift+Cmd+G) seçip şunu yazın:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   Klasör yoksa `Documents` içinde `wef` adıyla oluşturun.
2. Manifest dosyasını bu klasöre kopyalayın.
3. Word'ü tamamen kapatıp (Cmd+Q) yeniden açın.
4. **Ekle → Eklentilerim** → listede **Vefeyât (beta)** → ekleyin.

Mac'te kısayollar Alt+Shift yerine **Cmd+Option** ile çalışır (ör. Cmd+Option+V).

## 4. İlk deneme

1. [Deneme tezini](https://apoyraz58-max.github.io/vefeyat/beta/vefeyat-deneme.docx)
   indirip açın. İlk sayfasında adım adım bir kontrol listesi var.
2. Kendi metninizde: bir şahıs adını çift tıklayarak seçin → **Vefeyât (beta) → Ölüm
   tarihi ekle**. Ad tek ve kesinse tarih hemen eklenir; değilse sağdaki pencere seçim ister.
3. Bütün belge için **Şahıs dökümü** ya da **Tüm belgeyi denetle**.

Ayrıntılı kullanım: [kılavuz](https://github.com/apoyraz58-max/vefeyat/blob/main/docs/WORD-DENEME.md)
(sağdaki pencerede **Kullanım kılavuzu** düğmesiyle de açılır).

## 5. Güncellemeler

Pencerenin içindeki yenilikler kendiliğinden gelir; pencerenin en altında sürüm numarası
yazar (ör. `0.4.0-beta.12`). Şeride yeni bir düğme eklendiğinde size haber verilir; o zaman
manifest dosyasını yenisiyle değiştirip Word'ü kapatıp açmanız gerekir. Değiştirdikten
sonra şerit eski kalırsa Word kapalıyken şu klasörü silin:
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Mac'te: `~/Library/Containers/com.microsoft.Word/Data/Library/Caches`).

## 6. Sorun bildirmek

Sağdaki pencerenin en altındaki **Sorun bildir** bağlantısına tıklayın:

1. Ne yaptığınızı ve ne beklediğinizi bir iki cümleyle yazın.
2. **E-postayla gönder** yeni bir e-posta açar; ya da **Panoya kopyala** deyip kendi
   e-postanıza yapıştırın. Adres: **apoyraz@cumhuriyet.edu.tr**.

Bildirime eklenti ve Word sürümü, ayarlarınız ve o oturumdaki son mesajlar girer; **belge
metni girmez**. Sorunu gösteren kısa bir örnek cümle eklerseniz çözmek kolaylaşır.
Paneldeki mesajın ekran görüntüsü de çok işe yarar.

## 7. Kaldırmak

**Ekle → Eklentilerim** → Vefeyât (beta) üzerinde **…** → **Kaldır**. Ardından Güven
Merkezi'ndeki katalog yolunu silebilirsiniz. Vefeyât'ın eklediği tarihler belgede kalır;
teslimden önce **Teslim öncesi temizle** ile gizli Vefeyât işaretlerini kaldırabilirsiniz.
