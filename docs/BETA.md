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

## 1. Windows'ta kurulum (önerilen: tek tıkla)

1. [Vefeyat-Kurulum.zip](https://apoyraz58-max.github.io/vefeyat/beta/Vefeyat-Kurulum.zip)
   dosyasını indirin, sağ tıklayıp **Tümünü ayıkla** deyin (ya da açıp içindekileri
   masaüstüne sürükleyin).
2. Word'ü kapatın ve çıkan **Vefeyat-Kur.cmd** dosyasına **çift tıklayın**.
   - Windows "Windows bilgisayarınızı korudu" derse **Ek bilgi → Yine de çalıştır**'a basın.
     Dosya imzasız olduğu için bu uyarı çıkar; dosya yalnız aşağıdakileri yapar.
3. Siyah bir pencere açılır. Pencere Word sürümünüze uygun manifesti indirir, eklentiyi
   Word'e kaydeder ve "Tamam" yazar. Herhangi bir tuşa basıp pencereyi kapatın.
4. Word'ü açın. Şeritte **Vefeyât (beta)** sekmesi görünmezse (yalnız ilk seferde):
   **Giriş → Eklentiler** (bazı sürümlerde **Ekle → Eklentilerim**) → gerekirse
   **Diğer Eklentiler** → açılan pencerede **Geliştirici Eklentileri** bölümünde
   **Vefeyât (beta)** → **Ekle**. Sonraki açılışlarda Word eklentiyi kendisi yükler.

Dosya yönetici izni istemez; yalnız sizin kullanıcı hesabınıza kayıt yapar
(`HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer`) ve manifesti
`%LOCALAPPDATA%\Vefeyat` klasörüne koyar. Kaldırmak için aynı ZIP'teki
**Vefeyat-Kaldir.cmd**.

Kurulum olmazsa pencerenin ekran görüntüsünü apoyraz@cumhuriyet.edu.tr adresine gönderin
ya da aşağıdaki elle kurulumu deneyin.

## 2. Mac'te kurulum

**Uygulamalar → İzlenceler → Terminal**'i açın, aşağıdaki satırı yapıştırıp Enter'a basın,
sonra Word'ü tamamen kapatıp (Cmd+Q) yeniden açın:

```
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef && curl -fsSL https://apoyraz58-max.github.io/vefeyat/beta/manifest.xml -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/vefeyat-beta.xml && echo Tamam
```

Word'de **Ekle → Eklentilerim** → **Vefeyât (beta)**'yı ekleyin. Mac Word 16.88'den
eskiyse satırdaki `manifest.xml` yerine `manifest-uyumlu.xml` yazın. Mac'te kısayollar
Alt+Shift yerine **Cmd+Option** ile çalışır (ör. Cmd+Option+V).

## 3. Elle kurulum (Windows, kurulum dosyası çalışmazsa)

Word sürümünüze göre manifesti seçin:

| Word | Dosya |
|---|---|
| Microsoft 365, sürüm **2408 veya sonrası** | [manifest.xml](https://apoyraz58-max.github.io/vefeyat/beta/manifest.xml) |
| Daha eski Word, Office 2019/2021 | [manifest-uyumlu.xml](https://apoyraz58-max.github.io/vefeyat/beta/manifest-uyumlu.xml) |

Sürümünüzü görmek için: Word → **Dosya → Hesap → Word Hakkında**. İki manifest arasındaki
tek fark klavye kısayollarıdır; uyumlu sürümde bütün komutlar şeritte ve sağ tık menüsünde
aynen bulunur.

1. Bir klasör açın, ör. `C:\Vefeyat`, ve manifest dosyasını içine koyun.
2. Klasöre sağ tıklayın → **Özellikler → Paylaşım → Paylaş…** → kendinizi seçip
   **Paylaş**. Pencerede görünen ağ yolunu (`\\BILGISAYARADI\Vefeyat`) kopyalayın.
3. Word → **Dosya → Seçenekler → Güven Merkezi → Güven Merkezi Ayarları → Güvenilen
   Eklenti Katalogları**: yolu yapıştırın, **Katalog Ekle**, **Menüde göster**'i işaretleyin,
   **Tamam**.
4. Word'ü kapatıp açın → **Ekle → Eklentilerim → Paylaşılan Klasör → Vefeyât (beta)**.

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
**Vefeyat-Kur.cmd**'yi yeniden çalıştırın (Mac'te Terminal satırını yeniden yapıştırın;
elle kurduysanız manifest dosyasını yenisiyle değiştirin) ve Word'ü kapatıp açın. Değiştirdikten
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

Windows'ta **Vefeyat-Kaldir.cmd**'yi çalıştırın. Elle kurduysanız **Ekle → Eklentilerim** →
Vefeyât (beta) üzerinde **…** → **Kaldır**, ardından Güven Merkezi'ndeki katalog yolunu
silin. Mac'te `wef` klasöründeki `vefeyat-beta.xml` dosyasını silin. Vefeyât'ın eklediği tarihler belgede kalır;
teslimden önce **Teslim öncesi temizle** ile gizli Vefeyât işaretlerini kaldırabilirsiniz.
