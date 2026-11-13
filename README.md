# Vefeyât

Word için şahıs adı ve vefat tarihi denetimi. Tezdeki şahıs adlarını yerel DİA
kataloğuyla karşılaştırır, doğru yazımı önerir ve ölüm tarihini "(ö. 505/1111)"
biçiminde ekler. Belirsiz kimliklerde seçimi kullanıcıya bırakır; kaynakta
olmayan kişiye tarih uydurmaz; internete bağlanmaz.

## Depo yapısı

| Klasör | İçerik |
| --- | --- |
| `src/core/` | Word'den bağımsız çekirdek (TypeScript): normalleştirme, ad eşleştirme, CSV kataloğu, tarih biçimleri, seçili ad kuralları |
| `test/` | Çekirdek testleri; VBA eklentisinin yerleşik öz denetimleri birebir taşınmıştır |
| `legacy/vba-0.3.41/` | Dondurulmuş VBA eklentisi (son VBA sürümü), kurulum ve kabul belgeleri |
| `docs/` | Geçiş planı ve tasarım notları |

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

## Durum

- **VBA 0.3.41**: Windows Word için çalışan sürüm; yalnız hata düzeltmesi alır.
- **Office.js (0.4 geliştirme)**: 1. aşama tamam: eşleştirme çekirdeği ve testleri.
  Sonraki aşamalar `docs/MIGRATION.md` içinde.

Kişisel CSV kataloğu, tezler ve derlenmiş `.dotm` dosyaları depoya eklenmez.
