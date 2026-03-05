# Tarama & Barkod Okuma Uygulaması (Scanner & Barcode Reader)

Bu uygulama, masaüstü üzerinde belgeleri taramak, görüntüler üzerinden serbest kırpma yapmak, barkod okumak ve sonuçları Excel'e aktarmak için geliştirilmiş bir Python **Tkinter** uygulamasıdır.

## 🚀 Özellikler

- **TWAIN Tarayıcı Desteği**: Bağlı tarayıcılardan doğrudan uygulamanın içine tarama yapabilirsiniz.
- **Dosyadan Görüntü Açma**: Mevcut görüntü dosyalarınızı (`.png`, `.jpg`, `.jpeg`, vb.) uygulamaya yükleyebilirsiniz.
- **Serbest Kırpma**: Farenizle görüntü üzerinde dikdörtgen bir alan seçip ("ubber-band selection") kırpabilirsiniz.
- **Barkod Okuma**:
  - Taranan veya yüklenen görüntünün tamamından barkod okuyabilirsiniz.
  - Sadece kırptığınız (seçtiğiniz) alandan barkod okuyabilirsiniz.
  - Aynı görüntüde birden fazla barkod varsa uygulama hepsini algılar. En kısa olan barkodu öncelikli olarak seçer, ancak dilerseniz tümünü Excel tablosuna aktarabilirsiniz.
- **Excel'e Aktarma (Veri Tablosu)**:
  - Okunan barkodları sağdaki dinamik tabloya ekleyebilirsiniz.
  - Tabloya manuel satır/sütun ekleyebilir, sütun sıralarını sağa/sola kaydırarak değiştirebilirsiniz.
  - Tablodaki herhangi bir hücreye çift tıklayarak içeriğini düzenleyebilirsiniz.
  - Tabloyu şık bir `.xlsx` formatında dışa aktarabilirsiniz.
- **Görünüm Seçenekleri**:
  - Görüntüyü sağa/sola 90 derece veya tamamen 180 derece döndürebilme.
  - Görüntüyü yakınlaştırma/uzaklaştırma (Zoom in / Zoom out).

## 🛠️ Kurulum

Uygulamanın çalışması için sisteminizde Python 3 yüklü olmalıdır. Ardından gerekli kütüphaneleri kurmak için terminalinizde şu komutu çalıştırın:

```cmd
pip install Pillow pyzbar openpyxl pytwain
```

### Ek Gereksinimler (Sisteme Göre)
* **pyzbar** kütüphanesinin barkodları doğru çözebilmesi için işletim sisteminizde C++ derleyici (Visual C++) ve bazen ZBar sistem dll kütüphaneleri gerekebilir. Windows üzerinde genellikle `pip install pyzbar` kendi dll paketini indirir ancak hata alırsanız 64-bit bağımlılıklarını kurmayı ihmal etmeyin.
* Tarama işlevi cihazınızın TWAIN sürücüsünü kullanır. Tarayıcınızın kendi sürücülerinin kurulu olduğundan emin olun.

## 💻 Kullanım

Uygulamayı başlatmak için:

```cmd
python scanner_app.py
```

### 1- Görüntü Yükleme/Tarama
- Sol üstteki **📷 Tara** butonuna basarak tarayıcınızdan bir belge tarayabilirsiniz.
- Veya **📂 Dosya Aç** butonuna basarak bilgisayarınızdan bir görsel seçebilirsiniz.

### 2- Görüntü Düzenleme (Opsiyonel)
- Görüntü yüklendikten sonra sol taraftaki butonları kullanarak **180°**, **90° Sağa/Sola** döndürebilirsiniz.
- Farenizin sol tuşuna basılı tutarak görüntü üzerinde bir seçim alanı (kırpma alanı) çizebilirsiniz.

### 3- Barkod Okuma
- **🔍 Kırpılmış Görselden Oku**: Çizdiğiniz dikdörtgen alandaki barkodu analiz eder. Birden çok barkod varsa, metin uzunluğu en kısa olanı otomatik seçip tabloya aktarır (Diğerleri size ayrıca sunulacaktır).
- **🔍 Tüm Görselden Oku**: Resmin tamamındaki barkodları analiz eder.

### 4- Veri Tablosu (Sağ Panel)
- **Satır Ekle**: Boş bir satır ekler.
- **Kolon Ekle / Çıkar / Düzenle**: Verilerinize özel sütunlar yaratabilirsiniz. Sütunların başlığına tıklayarak sağındaki "<" (Sola Taşı) ">" (Sağa Taşı) veya "X" (Sil) seçeneklerini kullanabilirsiniz.
- **Hücre Düzenleme**: Bir satıra tıkladıktan sonra o satırdaki hücrelere çift tıklayarak serbest metin girebilirsiniz.
- **Excel'e Aktar**: Tüm tablo taslağını stilize şekilde bir Excel (`.xlsx`) dosyası olarak indirir.

## 🗂️ Dosya Yapısı

* `scanner_app.py`: Ana uygulama arayüzü, tarama sınıfları ve tablo işlemlerinin hepsini barındıran kaynak kod.
* `requirements.txt`: Bağımlılık paketlerinin listesidir (`pip install -r requirements.txt` şeklinde de kurulum yapılabilir).
