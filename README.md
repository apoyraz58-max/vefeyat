# Vefeyât

Word için şahıs adı ve vefat tarihi denetimi. Tezdeki şahıs adlarını yerel DİA
kataloğuyla karşılaştırır, doğru yazımı önerir ve ölüm tarihini "(ö. 505/1111)"
biçiminde ekler. Belirsiz kimliklerde seçimi kullanıcıya bırakır; kaynakta
olmayan kişiye tarih uydurmaz; internete bağlanmaz.

## Depo yapısı

| Klasör | İçerik |
| --- | --- |
| `src/addin/` | Office.js yan paneli: arayüz, Word işlemleri, katalog saklama |
| `src/core/` | Word'den bağımsız çekirdek (TypeScript): normalleştirme, ad eşleştirme, CSV kataloğu, tarih biçimleri, seçili ad kuralları |
| `test/` | Çekirdek testleri; VBA eklentisinin yerleşik öz denetimleri birebir taşınmıştır |
| `legacy/vba-0.3.41/` | Dondurulmuş VBA eklentisi (son VBA sürümü), kurulum ve kabul belgeleri |
| `manifest/`, `assets/` | Eklenti manifest şablonu; simgeler (Fluent UI System Icons, MIT, `assets/icon-src/NOTICE.txt`) |
| `data/` | Eklentiyle yayımlanan katalog (`catalog.csv`, herkese açık yayımlanır) |
| `scripts/` | Derleme (`build.mjs`) ve tarayıcıda duman testi (`smoke.mjs`) |
| `docs/` | Geçiş planı, Word'de deneme kılavuzu |

## Geliştirme

Node.js 20 veya üstü gerekir.

```bash
npm install
```

```bash
npm test
```

```bash
npm run typecheck
```

```bash
npm run build
```

```bash
npm run smoke
```

`npm run build` eklentiyi `dist/addin/` klasörüne derler (varsayılan adres GitHub
Pages; yerel sunucu için `BASE_URL=https://localhost:3000/ npm run build`).
`npm run smoke` derlenmiş paneli Chromium'da sahte bir Word belgesiyle baştan sona
çalıştırır.

## Durum

- **VBA 0.3.41**: Windows Word için çalışan sürüm; yalnız hata düzeltmesi alır.
- **Office.js (0.4 geliştirme)**: çekirdek (aşama 1) ve ilk yan panel (aşama 3) hazır:
  CSV kataloğu, seçili ad, aday seçimi, ölüm tarihi ekleme, "daha önce geçti" uyarısı,
  isim yazımı, kısaltma ayarı. Word'de denemek için `docs/WORD-DENEME.md`;
  sonraki aşamalar `docs/MIGRATION.md`.

Kişisel CSV kataloğu, tezler ve derlenmiş `.dotm` dosyaları depoya eklenmez.
