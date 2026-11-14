# VBA'dan Office.js'e geçiş planı

## Neden

VBA eklentisi yalnız Windows Word'de çalışır, makro izni ister, her sürümde
temiz kurulum gerektirir ve Word dışında test edilemez. Office.js eklentisi Mac,
Word Online ve iPad'de de çalışır, AppSource veya kurum yöneticisi üzerinden
kurulur ve mantığın büyük kısmı Word olmadan otomatik test edilebilir.

## İlke

Davranış korunur. Her taşınan fonksiyon VBA'daki karşılığının adını taşır ve
VBA'nın yerleşik öz denetimleri (`VfpRunPolicyChecks`, `VfaRunLookupChecks`)
TypeScript testlerine birebir aktarılır. VBA 0.3.41 dondurulur; yalnız hata
düzeltmesi alır.

## Aşamalar

| Aşama | Kapsam | Durum |
| --- | --- | --- |
| 1 | Çekirdek: normalleştirme, Türkçe ek ayıklama, ad puanlama, CSV kataloğu ve arama, kişi kimliği, tarih kısaltması ve algılama, seçili ad kuralları | Tamam |
| 2 | Metin taraması: paragraf metninde ad geçişlerini bulma (`VfbTextMatches`, `VfxClassify`), atıf ve bağlam kuralları; `academic-gold.json` ile ölçüm | Sırada |
| 3 | Office.js iskeleti: manifest, yan panel, katalog dosyası seçimi ve tarayıcıda saklama, "Ölüm tarihi ekle" ve "İsim yazımı" | Tamam (Word'de kabul bekliyor) |
| 4 | Belge işlemleri: içerik denetimleri (content control) ile tarih alanları, ilk geçiş denetimi, toplu önizleme, Şahıs dökümü | |
| 5 | Teslim araçları: tutarlılık raporu, şahıs raporu ve kronoloji, şahıslar dizini, teslim öncesi temizlik | |
| 6 | VBA belgelerinden geçiş: `ADDIN Vefeyat.Date.*` alanlarını ve belge değişkenlerini okuyup yeni biçime aktarma | |

## Taşınan fonksiyonlar (aşama 1)

| TypeScript | VBA |
| --- | --- |
| `normalize`, `isSeparator`, `collapseLetters`, `nameKey` | `VefeyatNormalize`, `VfIsSeparator`, `VfCollapseLetters`, `VfbKey` |
| `nameQuery`, `textName` | `VfNameQuery`, `VfnTextName` |
| `nameScore`, `nameMatch`, `looseNameMatch`, `spellingMatch`, `nearWord`, `prominence` | `VefeyatNameMatch`, `VfNameMatch`, `VfLooseNameMatch`, `VfSpellingMatch`, `VfNearWord`, `VfProminence` |
| `parseCatalogCsv`, `Catalog.find` | `VfParseCsv`, `VfAppendCsvRecord`, `VfValidateIdentities`, `VfBuildIndex`, `VefeyatFind` |
| `personKey`, `identity`, `articleUrl` | `VfpPersonKey`, `VfqIdentity`, `VfdArticleUrl` |
| `nameCandidates`, `confirmationCandidates`, `strength`, `singleSafe` | `VfaNameCandidates`, `VfaConfirmationCandidates`, `VfqStrength`, `VfqSingleSafe` |
| `formatDate`, `markerLength`, `requirePrefix`, `dateNote` | `VfpFormatDate`, `VfpMarkerLength`, `VfpRequirePrefix`, `VfqDateNote` |
| `startsWithDate`, `alreadyHasDate`, `dateKind`, `chronoYear` | `VfmDateText`, `VfaAlreadyHasDate`, `VfcDateKind`, `VfdChronoYear` |
| `selectionAlias`, `isValidAlias`, `isValidDate`, `aliasKey`, `orderBefore` | `VfaSelectionAlias`, `VfaValidAlias`, `VfaValidDate`, `VfaAliasKey`, `VfaOrderBefore` |

Bilinçli farklar: sorgu önbelleği (`mVfQueryKeys`) taşınmadı; arama saf
fonksiyondur. UTF-16 uzunlukları ve VBA `Trim$`/`LCase$` davranışı `text.ts`
içinde korunur ("İ" tek karakter kalır).

## Tasarım kararları (aşama 3–4 için açık)

- **Tarih alanı**: Office.js `ADDIN` alanı oluşturamaz. Aday: etiketli, gizli
  başlıklı içerik denetimi (`tag = "vefeyat:date:<id>"`). Zotero ile birlikte
  çalışma ve Word sürümlerinde görünüm doğrulanmalı.
- **Kimlik onayları ve ayarlar**: belge değişkenleri yerine `customXmlParts`
  veya belge ayarları (`document.settings`).
- **Katalog**: kullanıcı CSV'yi bir kez seçer; tarayıcı deposunda (IndexedDB)
  tutulur; belge metni hiçbir yere gönderilmez.
