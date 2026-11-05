> Bu belge önceki web sürümünün test kayıtlarıdır. Güncel çevrimdışı Word sürümü için [offline-validation.md](offline-validation.md) dosyasına bakın.

# İlk sürüm doğrulama kaydı

Bu kayıt, mevcut Linux bulut geliştirme makinesinde yapılan kontrolleri anlatır. İnternet yayını, başka bir görevde snapshot geri yükleme veya gerçek Windows Word kabul testi yapıldığı anlamına gelmez.

## Veri

- Kullanıcının Excel dosyasından 7.934 kayıt PostgreSQL'e aktarıldı.
- PostgreSQL'deki her kaydın altı özgün hücresi, kaynak satırı ve tarih metni Excel ile karşılaştırıldı; tamamı aynı.
- Aynı dosya tekrar aktarıldığında yeni kayıt eklenmedi ve mevcut kayıtlar değiştirilmedi.
- Kaynaklar bağımsız akademik doğrulamadan geçirilmedi; tarihlerin anlamı sessizce dönüştürülmedi.

## Backend

`python manage.py test catalog.tests --noinput`: 34 test başarılı, sıfır atlanan test. Gerçek PostgreSQL üzerinde ayrı test veritabanı kullanıldı.

İçe aktarımın atomikliği, aynı isimli kişilerin ayrımı, arama varyantları, yaklaşık eşleşme, tam/doğrudan eşleşmelerin önceliği, CSRF, geçersiz girdiler, katkıların beklemede kalması, editör yetkileri, kilitlenen kaydın güncel türüne göre yetki denetimi ve tek seferlik yayımlama doğrulandı.

`python manage.py check` ve `python manage.py makemigrations --check --dry-run`: sorun veya kaydedilmemiş model değişikliği yok.

## Word görev bölmesi

`node --test scripts/tests/word.test.mjs`: 15 Office/API taklit testi başarılı, sıfır atlanan test.

Geçerli seçimde yalnız özgün tarih ifadesinin eklenmesi, metin veya konum değiştiğinde yazmama, takip/karşılaştırma hatasında yazmama, WordApi 1.3 yoksa yazmanın kapalı olması, boş seçim, tekrar ekleme kontrolleri, manuel sorgunun seçimi geçersiz kılması, API hataları ve güvenli düz metin gösterimi doğrulandı.

HTTPS kök adresiyle üretilen manifest XML olarak ayrıştırıldı; kaynak adresleri ve WordApi 1.3 gereksinimi kontrol edildi. Resmi Office dağıtım servisi veya gerçek Windows Word üzerinde doğrulanmadı.

## Geliştirme servisleri

Servisler durdurulup `bash scripts/dev.sh` ile tekrar başlatıldı; veritabanı sorgusu yapan sağlık kontrolü geçti ve veri korundu. Aynı başlangıç adımları mevcut servisleri kullanabiliyor. PID/port sahipliği kontrolleri var; yabancı süreçler durdurulmuyor.

## Gerçek tarayıcı

`npx --no-install playwright test`: sistem Chromium'u ile 5 test başarılı, sıfır atlanan test. Masaüstü görünümünde gerçek veri, varyant arama, kaynaklı ayrıntı, aynı adlı iki şahıs, katkı penceresi, yaklaşık eşleşme, boş sonuç, mobil görünüm ve Word dışında manuel arama doğrulandı. Web arayüzünde ve 380px Word görev bölmesinde yatay taşma görülmedi.

Office köprüsü tarayıcı dışında kullanılmadığı için ilgili fallback testinde Office.js boş taklitle değiştirilir; veri araması gerçek API'ye gider. Bu test gerçek Word'ün yerini tutmaz.

## Dışarıda kalan kontroller

- Gerçek Windows masaüstü Word, WebView2, biçim koruma ve Geri Al davranışı.
- Windows'tan erişilen HTTPS adresi ve kullanıcı/kurum eklenti yükleme izinleri.
- Üretim dağıtımı, yedek geri yükleme ve yeni bulut görevinde ortam geri yükleme.
- Akademik pilot çalışma, SUS puanı ve zaman tasarrufu ölçümü.

Windows kabul adımları `docs/word-windows.md` içindedir.

## Windows yerel paketinin sunucu tarafı

Docker Compose yapılandırması ve başlangıç betiğinin kabuk sözdizimi doğrulandı. Ayrı Linux test projesinde PostgreSQL, nonroot uygulama ve Caddy başlatıldı; mevcut geliştirme servisleri değiştirilmedi. HTTPS üzerinden, yalnız bu testin kök sertifikasına güvenen bir TLS bağlamıyla sağlık, 7.934 kayıt, örnek arama, statik dosyalar ve manifest XML kontrol edildi. TLS doğrulaması kapatılmadı; makinenin genel sertifika deposu değiştirilmedi.

Kaynak Excel salt okunur bağlandı; imajda Excel veya yerel kimlik bilgisi bulunmadığı kontrol edildi. Yalnız yerel makinenin 8443 portu yayımlandı. Tekrar aktarım kayıt çoğaltmadı; manifest üretimi çalıştı.

HTTPS öneri gönderimi de doğrulandı: CSRF çerezi secure işaretli; eksik token 403, doğru token ve HTTPS Origin 201/pending döndürdü. Şahıs kaydı ve 7.934 kayıt sayısı değişmedi. Bu işlem yalnız ayrı test projesinde yapıldı; test projesi ve verileri temizlendi.

Bulut Docker build ağında DNS erişimi bulunmadığı için test imajı, host üzerinden indirilen ve requirements.txt SHA-256 özetleriyle doğrulanan aynı bağımlılık wheel'leriyle çevrimdışı kuruldu. Dağıtılan Dockerfile normal internet bağlantılı, checksum zorunlu pip kurulumunu kullanır. Windows'taki Docker Desktop ağının ve PowerShell akışının çalışması burada doğrulanmadı.

Paketleme için 8 güvenlik/içerik testi geçti. ZIP içindeki dosyaların özetleri doğrulandı; `.dockerignore` dahil gerekli kaynaklar var. Kullanıcının açıkça seçilen Excel dosyası dışında `.local` içeriği, parolalar, anahtarlar ve veritabanı pakete alınmadı.
