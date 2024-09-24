Excel VBA ile Taşımalı Araç Devam Takip İmza Çizelgesi Otomasyonu
Bu proje, Excel VBA kullanarak taşımalı araçlarla öğrenci taşımacılığı yapan okullar için günlük devam takip imza çizelgelerini otomatik olarak oluşturur. Araç plakaları, öğrenci sayıları ve imza alanları, belirlenen tarih aralığına göre çizelgede otomatik olarak eklenir. Proje, süreci hızlandırmak ve hata riskini en aza indirmek amacıyla geliştirilmiştir.

Özellikler
Otomatik Çizelge Oluşturma: Araç plakaları, öğrenci sayıları ve imza alanlarıyla birlikte, günlük devam takip çizelgeleri tarih aralığına göre otomatik oluşturulur.
Yazdırma ve Silme İşlevleri: Oluşturulan çizelgeler tek bir tıkla yazdırılabilir ve daha sonra kullanılmayan sayfalar otomatik olarak silinebilir.
Okul Müdürü ve Müdür Yardımcısı İsimleri: İlgili isimler Excel sayfasındaki E6 ve E7 hücrelerinden alınarak çizelgeye eklenir.
Esnek Hücre Koruma: Belirli hücreler şifre ile korunarak yanlışlıkla yapılan düzenlemeler önlenir.
Gereksinimler
Excel 2010 veya daha yeni bir sürüm
VBA Makro Güvenliği etkin olmalı (Makroları çalıştırmak için güvenlik ayarlarından makrolara izin verin)
Kurulum
Proje dosyasını indirin veya klonlayın:

bash
Kodu kopyala
git clone https://github.com/kullaniciadiniz/tasimali-arac-devam-takip-cizelgesi.git
Excel dosyasını açın ve Geliştirici Sekmesini etkinleştirin:

Dosya > Seçenekler > Şeritte Geliştirici'yi Göster yolunu takip edin.
Makroları etkinleştirin:

Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarlarına gidin ve Tüm Makroları Etkinleştir seçeneğini seçin.
Kullanım
Excel dosyasını açın.
Sayfa1'deki araç plakalarını ve öğrenci sayılarını girin.
Tarih aralıklarını ve okul bilgilerini doldurun.
Oluştur butonuna basarak çizelgeleri otomatik olarak oluşturun.
Çizelgeleri yazdırmak için "Yazdır" butonuna basın veya kullanılmayan sayfaları silmek için "Sil" butonunu kullanın.
Lisans
Bu proje MIT Lisansı altında lisanslanmıştır. Daha fazla bilgi için LICENSE dosyasına bakabilirsiniz
