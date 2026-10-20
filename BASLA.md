# Vefeyât ile ilk deneme

Bu paket, Vefeyât'ın ilk geliştirme sürümüdür. Kendi Windows bilgisayarınızda denemek için alan adı veya sunucu satın almanız gerekmez.

1. ZIP dosyasını bir klasöre çıkarın; içinde `BASLA.md` ve `compose.windows.yml` bulunan proje klasörünü açın.
2. Docker Desktop'ı kurup Linux containers modunda çalıştırın. İlk kurulumda internet bağlantısı gerekir.
3. Bu klasörde PowerShell açın. GitHub'dan indirilen sürüm Excel verilerinizi içermez. İlk başlatmada aşağıdaki komuttaki dosya yolunu kendi Excel dosyanızın konumuyla değiştirin:

   ```powershell
   .\scripts\windows-start.ps1 -DataFile "C:\Dosyalar\Şahıslar Veri Tabanı_251025.xlsx"
   ```

   Sonraki başlatmalarda `.\scripts\windows-start.ps1` komutu yeterlidir; yerel kopya `.local/source/people.xlsx` konumunda korunur.
4. [Yerel kurulum kılavuzunun](docs/windows-local.md) sertifika adımlarını uygulayın. Bu, Windows Word'ün kendi bilgisayarınızdaki uygulamaya güvenilir HTTPS ile bağlanmasını sağlar.
5. Önce web arayüzünde bir şahıs arayın. Ardından aynı kılavuzdaki Word manifestini indirme ve eklentiyi yükleme adımlarını izleyin.

İlk denemede Word belgesine **İbn Sînâ** yazıp adı seçin. Eklentide **Word'deki seçili adı al** düğmesine basın, doğru kaydı seçin ve tarih ifadesini ekleyin. Beklenen metin: **İbn Sînâ (ö. 428/1037)**.

**İbn Rüşd** araması iki farklı şahsı göstermelidir. Yaklaşık tarihler, yüzyıllar ve gaybet gibi ifadeler kaynakta bulunduğu biçimde korunur.

Web, veri aktarımı ve eklenti mantığı otomatik kontrollerden geçti. Gerçek Windows Word, sertifika kurulumu ve Word'ün Geri Al davranışı sizin bilgisayarınızda ayrıca denenmelidir. İnternet yayını yapılmamıştır.

Teknik ayrıntılar [README](README.md), yapılan kontroller [doğrulama kaydı](docs/validation.md) içindedir.
