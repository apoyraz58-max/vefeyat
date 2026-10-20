# Vefeyât

Akademik yazım için kaynaklı şahıs arama, biyografik tarih gösterimi ve Microsoft Word görev bölmesi eklentisi. İlk sürüm Windows masaüstü Word kullanımına göre hazırlanmıştır.

Windows'ta ilk deneme için [BASLA.md](BASLA.md) dosyasını izleyin. GitHub sürümü kaynak kodu ve kurulum dosyalarını içerir; Excel dosyanızı ilk başlatmada `-DataFile` parametresiyle kendiniz verirsiniz.

## İlk sürümün kapsamı

- Türkçe, Latin harfli ve orijinal ad alanlarında arama; isim varyantları ve yaklaşık eşleşme önerileri.
- Aynı isimli farklı şahıslar için ayrı sonuçlar, tarih, kısa tanıtım ve kaynak maddesi.
- Excel'den işlem bütünlüğü korunarak aktarım; ham hücreler, satır numarası ve dosya özeti saklanır.
- Kaynağın tarih ifadesini değiştirmeden kopyalama ve Word'deki seçili adın sonuna ekleme.
- Yeni şahıs/düzeltme önerileri ve yetkili editörün açık onayından sonra yayımlama.

Python 3.12+, Django 5.2 LTS ve PostgreSQL kullanılır. Web arayüzü sade JavaScript/CSS'tir; ön yüz derlemesi gerekmez. Node.js 20+ yalnız yerel PostgreSQL geliştirme servisi ve JavaScript kontrolleri için kullanılır.

## Bulut ortamında geliştirme

Gerekli araçlar: Python 3.12+, `uv`, Node.js 20+, `npm`. Mevcut checkout'u kullanın; her görev zaten yalıtılmıştır, ayrıca Git worktree oluşturmayın.

```sh
cd /workspace/vefeyat
bash scripts/dev.sh
```

Bu komut kilitli Python/Node bağımlılıklarını kurar, yalnız yerel makineye açık PostgreSQL'i başlatır, migration'ları uygular ve geliştirme web servisini 8000 portunda başlatır. PostgreSQL 55432 portunu kullanır. Yeniden çalıştırıldığında mevcut geliştirme süreçlerini korur; başka süreçlerin portlarını ele geçirmez.

```sh
# Yalnız web servisini yeni kodla yeniden başlatmak için:
node scripts/service-control.mjs restart-web

# Bu proje tarafından başlatılan servisleri durdurmak için:
bash scripts/stop.sh
```

Veri, günlükler ve yerel yapılandırma `.local/` altındadır. Servis süreçleri görevler arasında korunmuş varsayılmamalı; başlangıç betiği tekrar çalıştırılmalıdır. Yerel veritabanı parolası ve Django anahtarı otomatik üretilir, yalnız yerel dosyalarda saklanır ve Git'e alınmaz. Yayında ayrı anahtar ve veritabanı yapılandırması kullanın.

## Excel aktarımı

Kullanıcının gerçek Excel dosyası Git'e veya GitHub ZIP'ine eklenmez. Bu bulut geliştirme makinesinde korunan kaynak `.local/source/people.xlsx` konumundadır. Yeni bir kurulumda kendi Excel dosyanızı sağlayın; Windows başlangıç betiğinin `-DataFile` parametresi yerel kopyayı bu konuma oluşturur.

```sh
.venv/bin/python manage.py import_people .local/source/people.xlsx
```

Varsayılan sayfa `alimler_251025`; beklenen başlıklar sırasıyla `Kısa Ad`, `Tam Ad`, `Orijinal Ad`, `Vefat Tarihi`, `Tanıtım`, `Kaynak Link`.

- Aynı dosya özeti ikinci kez aktarıldığında işlem yapılmaz; editör değişiklikleri korunur.
- Yeni dosya mevcut kaynak URL'leriyle çakışıyorsa sessizce güncellemek yerine işlem durur.
- Hatalı satırda bütün aktarım geri alınır. Aynı isimli kişiler birleştirilmez.
- `---` gibi eksik değer işaretleri işlenmiş alanlarda boş bırakılır; özgün hücreler ayrıca saklanır.
- Tarihler ayrıştırılmaz ve otomatik takvim dönüşümü yapılmaz. `(1844-1921)`, `(m.ö. 384-322)`, `(gaybeti 260/874)` ve belirsizlik işaretleri aynen korunur.

Mevcut dosyadan 7.934 kayıt aktarılmıştır. Kaynak bağlantısı bulunması, tarihsel doğruluğun bağımsız teyit edildiği veya içerik için yeniden yayın izni alındığı anlamına gelmez. Toplu veri dışa aktarma uç noktası bu sürüme dahil değildir.

## Editör kullanımı

Kişisel editör hesabını etkileşimli olarak oluşturun; parolayı kaynak koduna veya sohbete yazmayın:

```sh
.venv/bin/python manage.py createsuperuser
```

Editör ekranı `/admin/` yolundadır. Katkılar önce incelemede bekler. Editör gerekli alanları düzenler, ardından “Seçili katkıları onayla ve yayımla” işlemini uygular. Ret işlemi şahıs verisini değiştirmez. Onay için katkı düzenleme ve şahıs düzenleme yetkisi, yeni şahısta ayrıca ekleme yetkisi gerekir. İncelenmiş katkının öneri alanları değiştirilemez; editör notu eklenebilir.

Bu sürümde kaynakların iki kişi tarafından akademik kontrol edildiği otomatik olarak varsayılmaz. İki bağımsız kaynak ve dört göz kontrolüyle ilgili araştırma protokolü ayrıca kararlaştırılmalıdır.

## Word eklentisi

Word görev bölmesi `/word/` yolundadır. Tarayıcıda manuel arama gösterilebilir; belgeye yazma yalnız Office içinde etkinleşir. Ayrıntılar: [Windows Word kurulumu](docs/word-windows.md).

Alan adı veya sunucu olmadan deneme için [Windows yerel kurulum](docs/windows-local.md) seçeneği hazırlanmıştır. Docker Desktop üzerinde yerel PostgreSQL ve HTTPS sunucusu çalıştırılır; geliştirme sertifikası kullanıcı tarafından güvenilir olarak tanıtılır. Bu, internet yayını gerektirmez.

Word için görev bölmesi, API ve ikonların güvenilir HTTPS adresinde erişilebilir olması gerekir. Bulut geliştirme makinesinin yerel adresi Windows'taki Word'e doğrudan ulaşmaz. HTTPS yayın adresi kesinleşince manifest üretin:

```sh
.venv/bin/python manage.py export_word_manifest \
  --base-url https://vefeyat.example.org \
  --output .local/vefeyat-manifest.xml
```

Bu örnek adresi kendi yayın adresinizle değiştirin. `WORD_BASE_URL` ayarı aynı adrese getirildiğinde `/word/manifest.xml` yolu da uygun manifesti sunar. Komut mevcut çıktı dosyasının üzerine yazmaz.

Gerçek Windows Word/WebView2 içinde seçim, ekleme, geri alma ve kurumsal eklenti izinleri kullanıcı ortamında ayrıca doğrulanmalıdır. Office API taklitleriyle geçen testler, gerçek Word uyumluluk testinin yerine geçmez.

## Kontroller

```sh
.venv/bin/python manage.py check
.venv/bin/python manage.py test catalog.tests --noinput
node --test scripts/tests/word.test.mjs
PLAYWRIGHT_BROWSERS_PATH=/workspace/.cache/ms-playwright npx playwright test
```

Tarayıcı testleri çalışan geliştirme servisine karşı yalnız okuma yapar ve 7.934 kayıtlı yerel veri setini bekler. Sistem Chromium'u varsa `/usr/bin/chromium` kullanılır; `CHROMIUM_PATH` ile başka bir kurulu yol seçilebilir. Sonuçlar `.local/test-results/` altına yazılır.

## API

| Yol | Davranış |
|---|---|
| `GET /api/health/` | Gerçek PostgreSQL sorgusu ile sağlık kontrolü |
| `GET /api/stats/` | Kayıt ve kaynak bağlantısı sayısı |
| `GET /api/people/?q=...&page=1&page_size=20` | Arama, yakın eşleşme işareti ve sayfalama (en fazla 50) |
| `GET /api/people/<uuid>/` | Şahıs ayrıntısı ve isim varyantları |
| `POST /api/contributions/` | CSRF korumalı, editör incelemesine giden öneri |

Öneri JSON alanları: `kind` (`new_person` / `correction`), düzeltmelerde `person_id`, `short_name`, `date_text`, `description`, `source_url`, `note`. `short_name`, HTTPS `source_url` ve `note` zorunludur. Düzeltmede boş tarih/tanıtım mevcut değeri silmez.

## Yayına hazırlık

Bu çalışma geliştirme sürümüdür; internet yayını yapılmadı. Yayında HTTPS ters vekili, kalıcı PostgreSQL ve yedekleme, statik dosya sunumu, kişisel editör hesapları ve katkı uç noktası için sunucu düzeyinde hız sınırı gerekir.

Çalışma ayarları: `DJANGO_DEBUG=0`, güçlü `DJANGO_SECRET_KEY`, `DATABASE_URL` (PostgreSQL; gerekirse `?sslmode=verify-full`), `DJANGO_ALLOWED_HOSTS`, `CSRF_TRUSTED_ORIGINS`, `WORD_BASE_URL`. Güvenilir ters vekil HTTPS başlığını kendisi ayarlıyorsa `TRUST_HTTPS_PROXY=1` kullanılabilir. Dış isteklerde gelen vekil başlıklarına doğrudan güvenmeyin.

```sh
.venv/bin/python manage.py collectstatic --noinput
.venv/bin/gunicorn vefeyat.wsgi:application --bind 127.0.0.1:8000
```

Ters vekil `/static/` yolunu `.local/static/` klasöründen sunmalıdır. Geliştirme sunucusu ve gömülü geliştirme PostgreSQL'i üretim dağıtımı olarak kullanmayın. Veri kullanım koşulları, alan adı ve gerçek Word pilot testi tamamlanmadan genel kullanıma hazır olduğu iddia edilmemelidir.
