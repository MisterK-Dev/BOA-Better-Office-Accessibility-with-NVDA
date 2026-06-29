# BOA: Better Office Accessibility

BOA, NVDA kullanıcıları için ekran okuyucu deneyimini büyük ölçüde geliştirmek üzere tasarlanmış, Microsoft Office için güçlü bir erişilebilirlik geliştirmeleri paketidir. Erişilemeyen kullanıcı arayüzü bileşenlerini doğrudan yamalar ve Excel ile PowerPoint için hızlı gezinti araçları sunar.

---

## ⌨️ Kısayol Başvurusu

| Özellik | Tuş Kombinasyonu | Bağlam / Notlar |
| :--- | :--- | :--- |
| **Komut Moduna Gir** | `[Prefix]` (Varsayılan: `NVDA+E`) | Komut Öneki Modunu etkinleştirir (tiz bir sinyal sesi tetikler) |
| **Komut Modunu İptal Et** | `Escape` | Komut Öneki Modundan çıkar |
| **EXCEL GELİŞTİRMELERİ** | | |
| **Sayfa Düzenini Çözümle** | `[Prefix]`, ardından `L` | Veri blokları arasında gezinmeden önce Excel içinde çalıştırın |
| **En Yakın Veri Bloğuna Atla** | `[Prefix]`, ardından `J` | Önce Düzen Çözümlemesi gerektirir |
| **Toplu Sayfa Düzenleyiciyi Aç** | `[Prefix]`, ardından `X` | Erişilebilir sayfa yeniden sıralama iletişim kutusunu açar |
| **Ham Formül Duyurucu** | `[Prefix]`, ardından `F2` | Ham formül dizesini duymak için tek dokunuş |
| **Güçlü Formül Düzenleyici** | `[Prefix]`, ardından iki kez `F2` | Erişilebilir çok satırlı formül düzenleyiciyi açmak için çift dokunuş |
| **Etkileyenleri İzle** | `[Prefix]`, ardından `Shift+P` | Etkileyenleri İzle özelliğinin aynısını erişilebilir bir şekilde sunar. |
| **Etkilenenleri İzle** | `[Prefix]`, ardından `Shift+D` | Etkilenenleri İzle özelliğinin aynısını erişilebilir bir şekilde sunar; bir hücre üzerinde Enter tuşuna basmak sizi o hücreye ışınlar. |
| **Ayrıntılı Koşullu Biçimlendirme** | `[Prefix]`, ardından `F` | Odaklanılan hücrenin tüm biçimlendirme ayrıntılarını duyurur |
| **Etkin Sayfayı Sola Taşı** | `NVDA+Shift+LeftArrow` | Etkin sayfayı bir konum yukarı kaydırır |
| **Etkin Sayfayı Sağa Taşı** | `NVDA+Shift+RightArrow` | Etkin çalışma sayfasını bir konum aşağı kaydırır |
| **Sayfayı Başlangıca/Sona Taşı** | `NVDA+Shift+Home` / `End` | Çalışma sayfasını mutlak sınırlara gönderir |
| **Satırı Gizle / Göster** | `Ctrl+9` / `Ctrl+Shift+9` | Yerel kısayol; BOA görünürlük değişikliğini açıkça duyurur |
| **Sütunu Gizle / Göster** | `Ctrl+0` / `Ctrl+Shift+0` | Yerel kısayol; BOA görünürlük değişikliğini açıkça duyurur |
| **Sütunu Göster (Alternatif)** | `NVDA+Ctrl+Shift+0` | Windows giriş dili kısayol çakışmalarını baypas eder |
| **Hücreyi Bellek Yuvasına Eşle** | `[Prefix]`, ardından `Shift+1` ila `Shift+9` | Geçerli hücreyi arka plan izleme yuvasına atar |
| **İzlenen Hücre Yuvasını Oku** | `[Prefix]`, ardından `1` ila `9` | Atanmış yuvanın değerini geri çağırır ve okur |
| **Doğrudan Yuvaya Atlama** | `Alt` + `1` ila `9` | İmlecinizi anında izlenen bir yuvaya atlatır |
| **Önceki Hücreye Geri Işınlan** | `[Prefix]`, ardından `\` | Bir yuvayı kontrol ettikten sonra sizi anında geri ışınlar |
| **Yuva Yöneticisi İletişim Kutusu** | `[Prefix]`, ardından `Alt+M` | Tüm etkin izleyicileri görüntülemek ve yönetmek için bir iletişim kutusu açar |
| **Arka Plan İzlemeyi Aç/Kapat** | `[Prefix]`, ardından `M` | Arka plan hesaplama takibini manuel olarak açıp kapatır |
| **Tüm Bellek Yuvalarını Temizle** | `[Prefix]`, ardından `Backspace` | Kaydedilmiş tüm arka plan hücre izleyicilerini temizler |
| **POWERPOINT GELİŞTİRMELERİ** | | |
| **Slayt Düzen Çözümleyici** | `[Prefix]`, ardından `L` | Geçerli slaydın uzamsal düzenini analiz eder ve duyurur |
| **Belge Çözümleyici** | `[Prefix]`, ardından `D` | Kapsamlı bir İçindekiler Tablosu ve sağlık raporu oluşturur |
| **Toplu Slayt Düzenleyici** | `[Prefix]`, ardından `X` | Birden çok slaydı yeniden sıralamak için erişilebilir iletişim kutusunu açar |
| **WORD GELİŞTİRMELERİ** | | |
| **Biçimlendirme Denetçisi** | `[Prefix]`, ardından `F` | Geçerli belgeyi biçimlendirme tutarsızlıkları açısından denetler |
| **Belge Çözümleyici** | `[Prefix]`, ardından `D` | Geçerli Word belgesinin düzenini ve yapısını analiz eder |

---

## 🚀 Özellikler

### Excel Geliştirmeleri

#### 1. Çalışma Sayfası Düzen Çözümleyici ve Önbelleğe Alma
Herhangi bir Excel çalışma sayfasını, yapısını, gizli öğelerini ve veri bloklarını anlamak için anında tarayın.
* **Nasıl çalışır:** BOA sayfayı hızlıca tarar ve etkin veri bloklarını duyurur. Ayrıca sizi **Gizli Çalışma Sayfası Sekmeleri**, etkin **Filtreler**, **Korumalı Modlar** ve **Gizli Dış Sınırlar** (örneğin, sayfanın sağ kenarına yakın sütunlar gizlenmişse, ekran dışındaki verileri kaçırmanızı önlemek amacıyla) hakkında uyarır.
* **Veri Gezintisi:** Taramadan sonra, bulunan veri blokları arasında imlecinizi anında ışınlamak için veri bloğu atlama kısayollarını kullanabilir, binlerce boş hücreyi zahmetsizce atlayabilirsiniz.

#### 2. Toplu Sayfa Düzenleyici
Tamamen erişilebilir bir iletişim kutusunu kullanarak birden fazla sayfayı aynı anda anında yeniden sıralayın ve düzenleyin.
* **Nasıl çalışır:** Bir sayfa seçip yeni bir konuma eşleyebileceğiniz bir iletişim kutusu açar. Planlanan taşımalar bir veri tablosunda listelenir (hatayı kaldırmak için `Del` tuşuna basın). `OK` düğmesine tıklayın ve çalışma kitabınız anında yeniden düzenlensin.

#### 3. Hızlı Sayfa Taşıyıcı
Etkin sayfayı klavye kısayollarınızı kullanarak sola, sağa, en başa veya en sona anında taşıyın.

#### 4. Erişilebilir Sayfa Yeniden Adlandırma
* Bir sayfayı yeniden adlandırırken, NVDA doğal olarak yazdığınız karakterleri okumakta zorlanır.
* BOA, `SafeRichEdit` motorunu kullanan özel bir `ExcelSheetRenameEdit` sınıfı enjekte eder; bu sayede yeniden adlandırırken karakter, sözcük veya satır bazında hassas bir şekilde okuma yapabilirsiniz. Bu özellik, mevcut varsayılan yeniden adlandırma davranışına bir geliştirme olarak hizmet eder.

#### 5. Gizli Satır/Sütun Takipçisi
* Gizli veya filtrelenmiş verileri kaçırmanızı önlemek için ızgara üzerindeki hareketinizi proaktif olarak takip eder.
* **Parçalanmış Hücreleri Geçme:** Izgaranın yoğun şekilde parçalanmış veya gizlenmiş bir bölümünün üzerinden atlarsanız (örneğin, Satır 4-9 gizlendiği için Satır 3'ten Satır 10'a geçmek), BOA açıkça "4 ile 9 arasındaki satırlar gizli" ("Rows 4 through 9 hidden") diye duyurur. Bu, yapıda ne zaman veri atlandığını her zaman bilmenizi sağlar.

#### 6. Koşullu Biçimlendirme Duyurucusu
* Excel'in Koşullu Biçimlendirme kuralları tarafından dinamik olarak değiştirilen hücrelerin rengini, yazı tipi stilini ve arka plan gölgesini otomatik olarak okur.
* Sadece alttaki ham değeri vermek yerine size hücrenin gerçek görsel durumunu sunar. Başlangıçta, hücreye odaklanıldığında "koşullu biçimlendirme var ve diğer bazı küçük ayrıntılar" şeklinde duyuru yapar. Kapsamlı bilgi için, `NVDA+E` ve `F` olan ayrıntılı kısayol yapılandırmasını kullanın.

#### 7. Daha iyi seçim duyurusu
hücre veya aralığın seçildiğini veya seçiminin kaldırıldığını okur.

#### 8 Hücre izleyici:
* **Hücre İzleyici:** Belirli hücreleri bellek yuvalarına eşlemek için komut yollarını kullanın. Atanmış sayısal yuvayı kullanarak istediğiniz zaman geri atlayabilir ve bunları okuyabilirsiniz.
* **Sürekli İzleme:** Yuvalara atanan hücreler arka planda otomatik olarak izlenir. Excel bir yeniden hesaplama or hücre düzenleme tetiklerse, BOA yeni değeri anında duyurur. Komut yuvaları aracılığıyla manuel olarak açıp kapatın veya tümünü temizleyin.
* **Excel: Hücre İzleyici Pro Yükseltmeleri:**
  - **Yuva Yöneticisi İletişim Kutusu (`NVDA+E`, ardından `Alt+M`):** Etkin olarak izlenen tüm hücrelerinizi listeleyen bir iletişim kutusu açar. Birine anında atlamak için `Enter` tuşuna basın.
  - **Geri Atla (`NVDA+E`, ardından `\`):** Bir yuvayı kontrol ettikten sonra sizi anında önceki çalışma hücrenize geri ışınlar.
  - **Doğrudan Yuvaya Atlama (`Prefix + Alt` + `Yuva Numarası`):** Öneki tamamen atlayıp atanmış bir hücre yuvasına anında atlayın.

#### 9 Güçlü Düzenleyici
* **Excel: Güçlü Düzenleyici (Erişilebilir Formül Düzenleyici):** Devasa formülleri değiştirmek için kuralları tamamen değiştiren bir özellik.
  - **Tek Dokunuş `NVDA+E`, ardından `F2`:** Etkin hücrenin ham formül dizesini anında duyurur (veya "Formül yok" diye duyurur).
  - **Çift Dokunuş `NVDA+E`, ardından `F2`:** Devasa, iç içe geçmiş formülleri güvenle değiştirmek için tam erişilebilir, çok satırlı bir düzenleyici açar. Yerel `Enter` kolay okuma için satır sonları ekler ve `Ctrl+Enter` bunu Excel'e geri kaydeder.
  - *Güvenlik Denetimleri:* Sayfanızı bozmadan önce sözdizimi hatalarını güvenle yakalar ve bir formül bozulduğunda sizi anında uyarmak için hesaplama sonrası hataları (`#NAME?` veya `#DIV/0!`) saptar.

#### 10 Formül denetimi ve değerlendirme geliştirmeleri:
* **Excel: Formül Denetimi ve Değerlendirme:** Etkileyenleri (Precedents) ve Etkilenenleri (Dependents) güvenilir bir şekilde izlemek için özel kısayollar (`NVDA+E`, ardından `Shift+P` ve `NVDA+E`, ardından `Shift+D`) eklendi. Ayrıca, Excel'in yerel "Formülü Değerlendir" iletişim kutusu artık tamamen erişilebilirdir; siz hesaplama adımlarında ilerlerken NVDA değerlendirilen sonuçları otomatik olarak okur!

### PowerPoint Geliştirmeleri

#### 1. Erişilebilir Renk Seçiciler
* PowerPoint'teki Özel Renk iletişim kutusunun kilidini açar.
* "Kırmızı", "Yeşil" ve "Mavi" düzenleme kutularını doğru bir şekilde tanımlar ve açıkça okur (`PowerPointRGBEdit` sınıfını geçersiz kılarak).
* Daha önce görünmeyen Hex giriş alanını eşler, böylece NVDA tam Hex renk değerini temiz bir şekilde okuyabilir.

#### 2. Standart Renk Izgarası Desteği
* PowerPoint "Standart" renk altıgen ızgarasında gezinmek normalde "Grafik" veya sessizlik olarak okunur.
* BOA, altıgen üzerindeki yön tuşlarınızı takip eder ve gizli renk değerini sessizce alıp size gerçek zamanlı olarak duyurur (örneğin, "Renk #FF0000").

#### 3 Toplu Slayt Düzenleyici:
* **PowerPoint: Toplu Slayt Düzenleyici (Deneysel) (`NVDA+E`, ardından `X`):** Excel özelliğine benzer şekilde, artık tamamen erişilebilir bir iletişim kutusu kullanarak birden fazla PowerPoint slaydını aynı anda anında yeniden sıralayabilir, taşıyabilir ve düzenleyebilirsiniz.

#### 4 Slayt düzen çözümleyici
* **PowerPoint: Slayt Düzen Çözümleyici (Deneysel) (`NVDA+E`, ardından `L`):** Tamamen sorunsuz ve duyarlı bir ekran okuyucu deneyimi sağlamak amacıyla, uzamsal düzenini ve erişilebilirlik kısıtlamalarını anlamak için şu anda etkin olan slaydınızı anında tarar. Yani burada, Excel'in sayfa düzen çözümleyicisine benzer şekilde geçerli slayt hakkında ayrıntılar alacaksınız.


#### 5 Tam Belge [PPT] Çözümleyici
* **PowerPoint: Tam Belge Çözümleyici (Deneysel) (`NVDA+E`, ardından `D`):** NVDA'nın konuşma motorunu dondurmadan tüm bir sununun haritasını çıkaran son derece gelişmiş, arka planda işlenen bir erişilebilirlik aracıdır. Derinlemesine gezilebilir bir Sanal İçindekiler Tablosu sağlar, Okuma Sırası Uyuşmazlıklarını (Görsel Sıraya Karşı Z Sırası) saptar, "Metin Duvarı" slaytlarını işaretler ve SmartArt ile Veri Tabloları gibi karmaşık nesnelerin haritasını çıkarır.

#### 6 Şekil hareketi [ayarlama] geliştirmeleri:
* **PowerPoint: Şekil Hareketi Ses Modu (Deneysel):** PowerPoint tuvaline 3B Uzamsal Ses ipuçları getirir. Bir nesneyi hareket ettirirken yönünü ve sınır limitlerini gösteren işitsel geri bildirim sağlayarak uzamsal farkındalığı büyük ölçüde artırır.

### Word Geliştirmeleri:
#### 1. Paul'un word access eklentisinden esinlenen ve türetilen Belge Çözümleyici:
* **Word: Belge Çözümleyici (`NVDA+E`, ardından `D`):** Word belgenizin yapısal genel bakışını anında ekrana getirin. *(Paul'e özel bir teşekkür ve kredi notu: Bu özellik, onun harika "Word Access" eklentisinden doğrudan ilham almıştır. Bu alandaki temel çalışması için ona derinden minnettarız!)*

#### 2 Biçimlendirme Denetçisi
* **Word: Biçimlendirme Denetçisi (`NVDA+E`, ardından `F`):** Görsel standartları sağlamak için Word belgenizi biçimlendirme tutarsızlıkları açısından denetler.

#### 3 Dipnot okuyucu:
* **Word: Otomatik Dipnot Duyurucusu:** Özel BOA ayarlarınıza bağlı olarak dipnotlar artık okurken satır içi olarak otomatik olarak duyurulacaktır. *(Not: Sonnotlar ve açıklamalar için destek gelecekteki bir sürümde planlanmaktadır).*

### Altyapı ve Teknik Mekanizmalar

#### Komut Öneki Modu
Diğer NVDA eklentileriyle tuş vuruşu çakışmalarını önlemek için BOA, bir **Komut Öneki Modu** kullanır:
1. Komut Moduna girmek için etkinleştirme kısayoluna basın. Tiz bir sinyal sesi duyacaksınız. Varsayılan değer NVDA artı E'dir.
2. Belirli bir özelliği tetiklemek için ikincil bir tuşa basın.
3. Geçersiz bir tuşa basarsanız bir hata sinyal sesi duyarsınız.

#### Özelleştirme ve Ayarlar Paneli
* BOA özellikleri tamamen modülerdir ve istendiği zaman etkinleştirilebilir veya devre dışı bırakılabilir. Ayrı özellikleri açıp kapatmak için `NVDA Menüsü -> Tercihler -> Ayarlar -> BOA Office Enhancements` yoluna gidin.
* **Akıllı Hızlandırıcı Tuşlar:** Paneldeki her bir ayar, matematiksel olarak benzersiz bir `Alt+Tuş` hızlandırıcı kısayoluna sahiptir. Örneğin, Excel grubuna anında geçmek için `Alt+E`, PowerPoint için `Alt+P` ve Word için `Alt+W` tuşlarına basın.
* Ayarlar, bağımsız bir JSON dosyasına (`boa_settings.json`) güvenli bir şekilde kaydedilir, böylece ana NVDA yapılandırmanızın asla bozulmaması sağlanır.
* Gelecekte Microsoft Office resmi olarak bir erişilebilirlik hatasını düzeltirse, eklentinin geri kalan işlevlerini kaybetmeden BOA'nın ilgili geçersiz kılma kancasını (override hook) güvenle devre dışı bırakabilirsiniz.
* **Girdi Hareketleri Özelleştirmesi:** Tüm Office uygulamalarındaki tüm özellikler, yerel NVDA Girdi Hareketleri iletişim kutusuna açıkça sunulmuştur; bu da size her klavye kısayolunu özelleştirme konusunda tam özgürlük sağlar.

#### Güvenlik ve Entegrasyon Sınırları
* Pano enjeksiyonları, diğer uygulamalara veri sızmasını önlemek amacıyla pencere ön plan işlem kimliklerini (PID) kesin bir şekilde doğrular.
* Bazı özel kısayol tuşları, NVDA'nın Girdi Hareketleri iletişim kutusunda "Better Office Accessibility" kategorisi altında tamamen sunulmuştur.

---

## 📋 Gereksinimler

* **NVDA:** Sürüm 2026.1.0 veya üzeri.
* **Uygulamalar:** Microsoft Excel ve Microsoft PowerPoint.

---

## 💾 Kurulum

1. En son `.nvda-addon` sürüm dosyasını indirin veya yerel NVDA Eklenti Mağazası'nda bulun.
2. Dosyadan yüklüyorsanız, dosyayı açın veya `NVDA Eklenti Mağazası -> Harici dosyadan yükle` seçeneğini kullanın.
3. NVDA'yı yeniden başlatın.

---

## 🛠️ Değişiklik Günlüğü

### Sürüm 2.0.0
#### Yeni Özellikler
* **PowerPoint: Tam Belge Çözümleyici (Deneysel) (`NVDA+E`, ardından `D`):** NVDA'nın konuşma motorunu dondurmadan tüm bir sununun haritasını çıkaran son derece gelişmiş, arka planda işlenen bir erişilebilirlik aracıdır. Derinlemesine gezilebilir bir Sanal İçindekiler Tablosu sağlar, Okuma Sırası Uyuşmazlıklarını (Görsel Sıraya Karşı Z Sırası) saptar, "Metin Duvarı" slaytlarını işaretler ve SmartArt ile Veri Tabloları gibi karmaşık nesnelerin haritasını çıkarır.
* **PowerPoint: Slayt Düzen Çözümleyici (Deneysel) (`NVDA+E`, ardından `L`):** Tamamen sorunsuz ve duyarlı bir ekran okuyucu deneyimi sağlamak amacıyla, uzamsal düzenini ve erişilebilirlik kısıtlamalarını anlamak için şu anda etkin olan slaydınızı anında tarar. Yani burada, Excel'in sayfa düzen çözümleyicisine benzer şekilde geçerli slayt hakkında ayrıntılar alacaksınız.
* **PowerPoint: Toplu Slayt Düzenleyici (Deneysel) (`NVDA+E`, ardından `X`):** Excel özelliğine benzer şekilde, artık tamamen erişilebilir bir iletişim kutusu kullanarak birden fazla PowerPoint slaydını aynı anda anında yeniden sıralayabilir, taşıyabilir ve düzenleyebilirsiniz.
* **PowerPoint: Şekil Hareketi Ses Modu (Deneysel):** PowerPoint tuvaline 3B Uzamsal Ses ipuçları getirir. Bir nesneyi hareket ettirirken yönünü ve sınır limitlerini gösteren işitsel geri bildirim sağlayarak uzamsal farkındalığı büyük ölçüde artırır. Belirtildiği gibi bu özellik deneyseldir; iyileştirilmesi için geri bildirimler beklenmektedir.
* **Word: Biçimlendirme Denetçisi (`NVDA+E`, ardından `F`):** Görsel standartları sağlamak için Word belgenizi biçimlendirme tutarsızlıkları açısından denetler.
* **Word: Belge Çözümleyici (`NVDA+E`, ardından `D`):** Word belgenizin yapısal genel bakışını anında ekrana getirin. *(Paul'e özel bir teşekkür ve kredi notu: Bu özellik, onun harika "Word Access" eklentisinden doğrudan ilham almıştır. Bu alandaki temel çalışması için ona derinden minnettarız!)*
* **Word: Otomatik Dipnot Duyurucusu:** Özel BOA ayarlarınıza bağlı olarak dipnotlar artık okurken satır içi olarak otomatik olarak duyurulacaktır. *(Not: Sonnotlar ve açıklamalar için destek gelecekteki bir sürümde planlanmaktadır).*
* **Excel: Güçlü Düzenleyici (Erişilebilir Formül Düzenleyici):** Devasa formülleri değiştirmek için kuralları tamamen değiştiren bir özellik.
  - **Tek Dokunuş `NVDA+E`, ardından `F2`:** Etkin hücrenin ham formül dizesini anında duyurur (veya "Formül yok" diye duyurur).
  - **Çift Dokunuş `NVDA+E`, ardından `F2`:** Devasa, iç içe geçmiş formülleri güvenle değiştirmek için tam erişilebilir, çok satırlı bir düzenleyici açar. Yerel `Enter` kolay okuma için satır sonları ekler ve `Ctrl+Enter` bunu Excel'e geri kaydeder.
  - *Güvenlik Denetimleri:* Sayfanızı bozmadan önce sözdizimi hatalarını güvenle yakalar ve bir formül bozulduğunda sizi anında uyarmak için hesaplama sonrası hataları (`#NAME?` veya `#DIV/0!`) saptar.
* **Excel: Formül Denetimi ve Değerlendirme:** Etkileyenleri (Precedents) ve Etkilenenleri (Dependents) güvenilir bir şekilde izlemek için özel kısayollar (`NVDA+E`, ardından `Shift+P` ve `NVDA+E`, ardından `Shift+D`) eklendi. Ayrıca, Excel'in yerel "Formülü Değerlendir" iletişim kutusu artık tamamen erişilebilirdir; siz hesaplama adımlarında ilerlerken NVDA değerlendirilen sonuçları otomatik olarak okur!
* **Excel: Hücre İzleyici Pro Yükseltmeleri:**
  - **Yuva Yöneticisi İletişim Kutusu (`NVDA+E`, ardından `Alt+M`):** Etkin olarak izlenen tüm hücrelerinizi listeleyen bir iletişim kutusu açar. Birine anında atlamak için `Enter` tuşuna basın.
  - **Geri Atla (`NVDA+E`, ardından `\`):** Bir yuvayı kontrol ettikten sonra sizi anında önceki çalışma hücrenize geri ışınlar.
  - **Doğrudan Yuvaya Atlama (`Alt` + `Yuva Numarası`):** Öneki tamamen atlayıp atanmış bir hücre yuvasına anında atlayın.
* **Girdi Hareketleri Özelleştirmesi:** Tüm Office uygulamalarındaki tüm özellikler, yerel NVDA Girdi Hareketleri iletişim kutusuna açıkça sunulmuştur; bu da size her klavye kısayolunu özelleştirme konusunda tam özgürlük sağlar.

#### UX/UI Geliştirmeleri
* **Birleşik Gözatılabilir Raporlar:** Eklenti genelinde birleşik bir HTML raporlama sistemi benimsedik. Excel Koşullu Biçimlendirme Duyurucusu, Düzen Çözümleyiciler ve Belge Çözümleyiciler gibi özellikler artık sadece devasa metin blokları konuşmuyor; sonuçları artık verileri kendi hızınızda incelemenize olanak tanıyan yerel, gezilebilir bir HTML penceresinde açılıyor.
* **Excel: Geliştirilmiş Etkilenenler/Etkileyenler İzlemesi:** Excel'in yerel formül izleme kısayolları için konuşma çıktısı büyük ölçüde geliştirildi (Doğrudan Etkileyenler için `Ctrl+[` ve Doğrudan Etkilenenler için `Ctrl+]`). NVDA artık tam olarak hangi hücrelerin seçildiğini açıkça duyuracaktır.
* **Excel: Birleştirilmiş Hücre Desteği:** Birleştirilmiş hücreler artık boşluk atlayan hücre izleyici tarafından doğru bir şekilde saptanıyor ve açıkça duyuruluyor.

#### Hata Düzeltmeleri
* **Word: Liste Öğelerinin Çift Okunması:** Belirli Word görünümlerinde NVDA'nın paragraf liste öğelerini çift okuması hatasını düzeltmek için geçici bir yama uygulandı.
* **Excel: Hücre İzleyici Yerelleştirme Hatası:** Son çeviri yerelleştirme güncellemelerinin neden olduğu altta yatan izleme hataları giderildi.

### v1.6.1 sürümündeki Yenilikler
* **Derin Dosya Yerelleştirmesi**: %100 yerelleştirme kapsamı sağlamak için Excel geliştirme modüllerinin (Sayfa Düzeni Çözümleyici ve Hızlı Sayfa Taşıyıcı gibi) derinlerindeki eksik dize çevirileri düzeltildi.
* **Genişletilmiş Çeviri Desteği**: Sisteme 7 yeni dil eklendi (Türkçe, Lehçe, Korece, Ukraynaca, Çekçe, Urduca ve Pencapça).
  *(Not: Bu çeviriler yapay zeka tarafından oluşturulduğundan bazı küçük çeviri hataları veya yanlışlıklar olabilir.)*

### v1.6.0
* **Kapsamlı Çeviri Desteği**: Eklenti artık 17 küresel dil desteğiyle tamamen yerelleştirildi.
  *(Not: Bu çeviriler yapay zeka tarafından oluşturulduğundan bazı küçük çeviri hataları veya yanlışlıklar olabilir.)*
* **Sıkı Kod Yönetimi**: Tüm kod tabanında GPL-2.0 telif hakkı üst bilgileri uygulandı."""),

### Sürüm 1.5.0
#### Yeni Özellikler
##### Veri Sonu Radarı
Büyük elektronik tablolarda gezinirken, boş bir hücrenin listenin sonuna geldiğiniz anlamına mı geldiğini yoksa verilerde sadece bir boşluk mu olduğunu söylemek zor olabilir. **Veri Sonu Radarı**, sizi boşlukta körüöküne ilerlemekten kurtarmak için akıllı bir çevre kontrolü görevi görür.
Boş bir hücreye her girdiğinizde BOA, hareket yönünüzdeki kalan hücreleri anında tarar. Kesinlikle hiç veri kalmamışsa, proaktif olarak şu duyuruları yapar:
* *"Aşağıda başka veri yok" ("No more data below")*
* *"Yukarıda başka veri yok" ("No more data above")*
* *"Sağda başka veri yok" ("No more data to the right")*
* *"Solda başka veri yok" ("No more data to the left")*
**Yapılandırma Seçenekleri:**
Bu özelliği `NVDA Tercihleri -> Ayarlar -> BOA Office Enhancements` aracılığıyla yapılandırabilirsiniz. Elektronik tablolar gizli karmaşıklıklar içerebileceğinden (görünmez formüller veya daraltılmış satırlar gibi), radar üç çalışma modu sunar:
1. **Kapalı (Off)**: Radarı tamamen devre dışı bırakır.
2. **Katı Bellek Denetimi (CountA) [Varsayılan]**: En güvenli ve en hızlı yaklaşım. Elektronik tablonun ham belleğini kontrol eder. Altınızda herhangi bir şey saptarsa (gizli satırlar, metin, sayılar veya görünmez formüller dahil), yanlış alarmları önlemek için tamamen sessiz kalır. Yalnızca sayfanın geri kalanı matematiksel olarak %100 boş olduğunda "Başka veri yok" duyurusunu yapar.
3. **Yalnızca Görünür Veriler (Matematik Motoru)**: Karmaşık sayfalar için tasarlanmış son derece gelişmiş bir motor. Gizli satırları ve görünmez formülleri (örneğin, `=""`) akıllıca filtreler. Yalnızca yolunuzda gerçek, görünür sayılar veya metinler kalmışsa sessiz kalır.

### Sürüm 1.4 - 2026-06-12
#### Yeni Özellikler
* **Hücre İzleyici:** Belirli hücreleri bellek yuvalarına eşlemek için komut yollarını kullanın. Atanmış sayısal yuvayı kullanarak istediğiniz zaman geri atlayabilir ve bunları okuyabilirsiniz.
* **Sürekli İzleme:** Yuvalara atanan hücreler arka planda otomatik olarak izlenir. Excel bir yeniden hesaplama veya hücre düzenleme tetiklerse, BOA yeni değeri anında duyurur. Komut yuvaları aracılığıyla manuel olarak açıp kapatın veya tümünü temizleyin.

#### Hata Düzeltmeleri

### Sürüm 1.3.0 — 2026-06-05
*Final sürümü.*

#### Yeni Özellikler
* **Çalışma Sayfası Düzen Çözümleyici:** Güçlü düzen tarama altyapısı eklendi. Sayfa Korumasını, etkin Sütun Filtrelerini, Gizli Çalışma Sayfası Sekmelerini ve gizli mutlak sınırları anında saptarken, bulunan veri bloklarını önbelleğe alır.
* **Kılavuzlu Veri Bloğu Gezintisi:** Çözümleme sonrası gezinti, boş hücreleri sorunsuz bir şekilde atlayarak imlecin ana veri kümeleri arasında anında ışınlanmasını sağlar.
* **Koşullu Biçimlendirme Duyurucusu:** Excel'in Koşullu Biçimlendirme kuralları tarafından değiştirilen hücrelerin dinamik rengini, yazı tipi stilini ve arka plan gölgesini otomatik olarak saptar ve okur.
* **Belirgin Ayar Hızlandırıcıları:** NVDA mimarisine kesin olarak uymak için BOA Ayarları GUI'si tamamen elden geçirildi. Her özellik onay kutusu artık küresel olarak benzersiz bir `Alt+Harf` kısayoluna sahiptir, bu da klavye döngüsünü önler ve ilk harfle gezinti hatalarını ortadan kaldırır.

#### Hata Düzeltmeleri
* **Mutlak Kenar Sınırı Algılama:** Etkin veri bloğunun çok dışında kalsalar bile gizli satırların/sütunların saptanmasını garanti etmek için yerel COM `UsedRange` kenar kontrolleri yerine mutlak 1D matematiksel sınır kontrolleri (`Row 1048576` ve `Column 16384`) getirildi.
* **Lazy COM Property Safe Bailouts:** Milyonlarca bitişik gizli yapıyı değerlendirirken NVDA iş parçacığının dondurulmasını önlemek için COM özellik döngüleri güçlendirildi.

### Sürüm 1.2.0 — 2026-06-03
*Final sürümü.*

#### Yeni Özellikler
* **App-Launch Caching:** Büyük mimari revizyon. Çekirdek modüller artık tam olarak Office uygulamalarına odaklandığınızda geç yüklenir (lazy-load), bu da başlatma gecikmesini ortadan kaldırır, yeniden adlandırma iletişim kutularındaki 'bilinmeyen' (unknown) nesne odak hatasını tamamen çözer ve çoklu dosya kod tabanı yapısını korur.
* **Geliştirilmiş Hücre Takipçisi (1D COM Matematiği):** Gizli hücre boşluğu algılama mantığı, yalnızca tek boyutlu kesitleri (`current_col` veya `current_row`) değerlendirecek şekilde yeniden yazıldı. Bu, COM hesaplama yükünü 16 milyondan fazla hücre kadar azaltarak gizli aralıkları atlarken oluşan gezinti donmalarını anında ortadan kaldırır.
* **İşlem Belleği Temizleme:** Kullanıcının Excel'i ne zaman kapatıp yeniden açtığını saptamak için Excel Pencere Tanıtıcısı (`Hwnd`) takibi uygulandı. Bu, eski genel durum belleklerini aktif olarak siler ve yeni bir "Book1" açıldığında oluşan yanlış "Sayfa gizli" duyurusunu tamamen çözer.

#### Hata Düzeltmeleri
* **Çift Seçim Duyurusu:** Güvenilmez eşzamansız `winUser.getKeyState` kullanımından vazgeçilerek Shift+Yön tuşları kullanılırken çift duyuruları mükemmel şekilde bastırmak için `api.getLastInputGesture()` uygulandı.
* **Sınır Algılayıcıyı Devre Dışı Bırakma:** Proaktif Sınır Algılayıcı, NVDA yerel gezinti kararlılığını korumak için devre dışı bırakıldı ve tamamen boşluk atlayan takipçiye geri dönüldü.

### Sürüm 1.1.0 — 2026-05-30
*Final sürümü.*

#### Yeni Özellikler
* **Ayarlar GUI'si:** Özellikleri kolayca açıp kapatmak için `NVDA -> Tercihler -> Ayarlar` içine yerel bir BOA Office Enhancements paneli eklendi.
* **SafeRichEdit Kancası:** Office 2024'teki RichEdit kontrolleriyle etkileşim kurarken sessiz NVDA çökmelerini önler.
* **Özelleştirilebilir Kısayollar:** Tüm BOA kısayol tuşları artık NVDA'nın Girdi Hareketleri iletişim kutusunda "Better Office Accessibility" kategorisi altında tamamen sunulmuştur.
* **Excel: Gizli Satır/Sütun Atlamayı Algılama:** Gizli satırları veya sütunları geçerken proaktif olarak duyuru yaparak filtrelenmiş verileri asla kaçırmamanızı sağlar. Ayarlardan açılıp kapatılabilir.

#### Hata Düzeltmeleri
* **İş Parçacığı Güvenliği (Thread Safety):** Tüm engelleyici gecikmeler (`time.sleep`) kaldırıldı ve arka plan işlemleri sırasında ekran okuyucunun asla teklememesini sağlamak amacıyla engelleyici olmayan NVDA eşzamansız geri çağırmalarıyla değiştirildi.

### Sürüm 1.0.0 — 2026-05-24
*İlk genel sürüm.*

#### Yeni Özellikler
* **Excel: Toplu Sayfa Düzenleyici:** Tamamen erişilebilir bir iletişim kutunu kullanarak birden fazla sayfayı aynı anda anında yeniden sıralayın.
* **Excel: Hızlı Sayfa Taşıyıcı:** Klavye komutlarıyla etkin sayfayı sola, sağa, başlangıca veya sona taşıyın.
* **Excel: Erişilebilir Sayfa Yeniden Adlandırma:** Erişilemeyen yerel yeniden adlandırma alanını yakalar ve bunu güvenilir, erişilebilir bir iletişim kutusuyla değiştirir.
* **Excel: Akıllı Seçim Takibi:** Çoklu hücre aralığı seçimlerini ve seçim kaldırmalarını doğru bir şekilde duyurur.
* **PowerPoint: Erişilebilir Renk Seçiciler:** NVDA'nın Özel Renk iletişim kutusundaki RGB ve Hex değerlerini doğru bir şekilde okumasını sağlar.
* **PowerPoint: Standart Renk Izgarası Desteği:** Erişilemeyen renk altıgen ızgarasındaki gizli Hex kodlarını okumak için yön tuşu gezintisini yakalar.
