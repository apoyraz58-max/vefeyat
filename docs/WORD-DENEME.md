# Office.js sürümünü Word'de deneme (0.4 geliştirme)

Bu sürüm VBA eklentisinden bağımsızdır; ikisi aynı anda kurulu olabilir.
Önce bir tez **kopyasında** deneyin.

## 1. Yayını açın (bir kez)

GitHub'da depo → **Settings → Pages → Build and deployment → Source: GitHub
Actions**. `main` dalına her birleştirmede eklenti
`https://apoyraz58-max.github.io/vefeyat/` adresine yayımlanır. Manifest dosyası
`https://apoyraz58-max.github.io/vefeyat/manifest.xml` adresindedir.

## 2. Manifesti Word'e yükleyin

**Word Online (en kolay):** Bir belge açın → **Giriş → Eklentiler** (veya
**Ekle → Eklentiler**) → **Eklentilerim → Eklentimi Karşıya Yükle** →
`manifest.xml` dosyasını seçin.

**Windows masaüstü Word:** Manifesti bir klasöre koyun ve klasörü paylaşın
(sağ tık → Özellikler → Paylaşım). Word → **Dosya → Seçenekler → Güven Merkezi
→ Güven Merkezi Ayarları → Güvenilen Eklenti Katalogları** → paylaşılan klasörün
`\\BILGISAYAR\klasor` yolunu ekleyin, "Menüde göster"i işaretleyin, Word'ü
yeniden başlatın → **Ekle → Eklentilerim → Paylaşılan Klasör → Vefeyât**.

**Mac Word:** Manifesti
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef` klasörüne kopyalayın
(klasör yoksa oluşturun), Word'ü yeniden başlatın → **Ekle → Eklentilerim**.

## 3. Kullanım

1. **Giriş** sekmesinde **Vefeyât** düğmesi paneli açar.
2. **CSV seç** ile DİA kataloğunuzu bir kez seçin. Dosya yalnız bu bilgisayarda,
   eklentinin kendi tarayıcı deposunda tutulur.
3. Metinde bir adı seçin. Tek ve güçlü eşleşmede şahıs kendiliğinden seçilir;
   aynı adlı birden fazla şahıs veya benzer yazımda listeden seçin.
4. **Ölüm tarihi ekle**: Ad metinde daha önce geçiyorsa uyarı çıkar
   (**Önceki geçişe git** / **Yine de ekle**). Adın arkasında zaten tarih varsa
   eklenmez.
5. **İsim yazımı**: Önerilen yazımı seçip **Değiştir**. Türkçe ek korunur
   ("Gazali'nin" → "Gazzâlî'nin"). Yalnız adı seçin.
6. **Ayarlar → Kısaltma**: ö. / öl. / v. / kaynakta olduğu gibi. Belgeyle
   birlikte kaydedilir.

## Bu sürümün sınırları

- Tarih, gizli görünümlü ve `vefeyat:date:…` etiketli bir içerik denetimine
  yazılır (VBA sürümündeki gibi Word alanı değil). Zotero ile birlikte nasıl
  davrandığını not edin.
- "Daha önce geçti" denetimi yalnız ana metinde arar; dipnotlar henüz yok.
  Aynı yazılışı arar; farklı transkripsiyonla yazılmış önceki geçişi bulamaz.
- Şahıs dökümü, toplu denetim, dizin, raporlar ve teslim temizliği henüz yok
  (VBA sürümünde var).

## Geri bildirim

Bir sorun görürseniz: Word sürümü (Online / Windows / Mac), ne yaptığınız ve
paneldeki mesajın ekran görüntüsü. Hata mesajlarının sonunda sürüm yazar.
