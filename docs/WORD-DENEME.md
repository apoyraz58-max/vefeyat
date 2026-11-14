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

**Şerit.** Üstte **Vefeyât** sekmesi, VBA sürümündeki gibi beş grup:

| Grup | Düğme | Ne yapar |
| --- | --- | --- |
| Ekle | **Ölüm tarihi ekle** | Seçili ad tek ve kesin eşleşiyorsa, ilk geçişse ve arkasında tarih yoksa tarihi hemen ekler; aksi hâlde Vefeyât penceresini açıp seçim veya onay ister. |
| Ekle | **İsim yazımı** | Yazım önerilerini pencerede gösterir. |
| Belge | **Şahıs dökümü**, **Belgeyi denetle** | (taşınıyor) |
| Teslim | **Teslim araçları ▾** | Tutarlılık raporu, Şahıs raporu ve kronoloji, Şahıslar dizini, Teslim öncesi temizle (taşınıyor) |
| Tercihler | **Kısaltma ▾**, **Tarih ekleme ▾**, **Ad biçimi ▾**, **Denetlenecek ▾** | Seçince hemen bu belgeye kaydedilir. Word, Office.js eklentilerinde seçili seçeneği işaretlemeye izin vermediği için geçerli değerler Vefeyât penceresinin üstünde yazar. |
| Veri ve yardım | **Katalog**, **Kılavuz** | Yüklü kataloğu gösterir / değiştirir; kısa kılavuz. |

**Sağ tık.** Metinde bir adı seçip sağ tıklayın → **Vefeyât → Ölüm tarihi ekle / İsim yazımı**.

**Pencere.** Sağdaki Vefeyât penceresi, VBA'daki pencereler gibi yalnız düğmenin istediği işi
gösterir (aday seçimi, onay, katalog, kılavuz). Ayar içermez.

**Katalog.** Eklentiyle yayımlanan DİA kataloğu kendiliğinden yüklenir. **Katalog → Başka bir
CSV seç** ile kendi dosyanızı kullanabilirsiniz; **Eklentinin kataloğuna dön** ile geri
dönersiniz.

**Tarih ekleme kuralları.** "Yalnız ilk kullanım" (varsayılan): ad metinde daha önce
geçiyorsa uyarı çıkar (**Önceki geçişe git** / **Yine de ekle**). "Bütün kullanımlar": uyarı
çıkmaz. "Yalnız denetle": tarih eklenmez. Her durumda adın arkasında zaten tarih varsa
ikinci tarih eklenmez.

Panel açıkken şerit düğmesine yeniden basıldığında Word pencereyi yeniden yüklemeyebilir;
bu durumda penceredeki düğmeyi kullanın.

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
