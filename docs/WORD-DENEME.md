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

**Şerit.** Üstte **Vefeyât** sekmesi, VBA 0.3.41 ile aynı yedi grup:

| Grup | Komutlar |
| --- | --- |
| Ekle ve denetle | **Ölüm tarihi ekle**, **Tüm belgeyi denetle**, **İsim yazımı** |
| Denetim | **Tarih ekleme ▾**, **Denetlenecek ▾**, **Denetime devam et**, **Toplu düzeltme önizlemesi** |
| Şahıs ve kaynak | **Şahıs ara**, **DİA maddesini aç**, **Tarihin kaynağı** |
| Belge | **Şahıs dökümü**, **Diğer araçlar ▾**, **Belge sağlığı**, **İşlem geçmişi** |
| Teslim ve dizin | **Tutarlılık raporu**, **Şahıs raporu** (yeni belge açar), **Şahıslar dizini ▾**, **Teslim öncesi temizle** |
| Tercihler | **Ad biçimi ▾**, **Kısaltma ▾**, **Ayarlar ▾** (yazım tercihi, tercihleri dışa aktar/al, tarih alanlarını göster/gizle) |
| Veri ve yardım | **Veri bilgisi**, **CSV seç**, **Kullanım kılavuzu** |

VBA sekmesindeki bütün komutlar bu sürümde çalışıyor; VBA eklentisini kaldırabilirsiniz
(önce aşağıdaki "VBA belgeleri" notuna bakın). VBA'daki değer gösteren açılır kutular (Tarih ekleme,
Denetlenecek, Ad biçimi, Kısaltma) Word'ün Office.js eklentilerine izin vermediği için menüdür;
seçili değerler Vefeyât penceresinin üstünde yazar.

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
- "Daha önce geçti" denetimi dipnotları WordApi 1.5 olan Word'de arar.
  Katalogdaki yazılışları ve seçili yazılışı arar; bunların dışında bir
  transkripsiyonla yazılmış önceki geçişi bulamaz.
- Sonnotlar, üst/alt bilgiler, açıklamalar ve metin kutuları taranmaz
  (Diğer araçlar → Denetim kapsamını göster).
- **VBA belgeleri:** VBA sürümünün eklediği tarihler Word alanıdır
  (`ADDIN Vefeyat.Date`). Yeni sürüm bunları metin olarak görür; Belge sağlığı,
  Tarihin kaynağı ve Teslim öncesi temizle bunları tanımaz. VBA ile başlanmış
  bir belgeyi teslimden önce VBA'daki Teslim öncesi temizle ile bir kez
  temizleyin ya da VBA eklentisini o belge bitene kadar tutun.
- Şahıs raporu yeni belge açar; Word Online'da açılır pencere engelleyicisi
  buna izin vermeyebilir.

## Geri bildirim

Bir sorun görürseniz: Word sürümü (Online / Windows / Mac), ne yaptığınız ve
paneldeki mesajın ekran görüntüsü. Hata mesajlarının sonunda sürüm yazar.
