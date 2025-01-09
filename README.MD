# XLSX'ten XLS'e Dönüştürme Uygulaması

Bu uygulama, XLSX formatındaki Excel dosyalarını eski format olan XLS formatına dönüştürmek için geliştirilmiş bir C# konsol uygulamasıdır.

## Özellikler

- XLSX dosyalarını otomatik olarak XLS formatına dönüştürme
- Tüm hücreleri metin formatına çevirme
- Toplu dönüştürme işlemi
- Basit ve kullanıcı dostu arayüz

## Gereksinimler

- Windows işletim sistemi
- .NET Framework 4.5 veya üzeri
- Microsoft Office Excel (bilgisayarda yüklü olmalı)
- Microsoft.Office.Interop.Excel referansı

## Kurulum

1. Projeyi bilgisayarınıza indirin
2. Uygulamanın bulunduğu dizinde `listeler` ve `bitenler` adında iki klasör oluşturulacaktır
3. Uygulamayı çalıştırın

## Kullanım

1. Dönüştürmek istediğiniz XLSX dosyalarını `listeler` klasörüne kopyalayın
2. Uygulamayı çalıştırın
3. Dönüştürülen XLS dosyaları `bitenler` klasöründe oluşturulacaktır
4. İşlem tamamlandığında uygulama sizi bilgilendirecektir

## Klasör Yapısı

XLSX_to_XLS/
│
├── Program.exe
├── listeler/ # XLSX dosyalarının konulacağı klasör
└── bitenler/ # Dönüştürülen XLS dosyalarının kaydedileceği klasör

## Hata Durumları

Uygulama aşağıdaki durumlarda hata mesajı gösterebilir:
- Excel uygulamasına erişilemediğinde
- Dosya okuma/yazma izinleri yetersiz olduğunda
- Kaynak dosya bozuk olduğunda
- Hedef dosya başka bir uygulama tarafından kullanıldığında

## Güvenlik Notları

- Uygulama, Microsoft Office Interop kullanarak Excel dosyalarını işler
- Dönüştürme işlemi sırasında orijinal dosyalar değiştirilmez
- Her işlem sonrası Excel uygulaması düzgün şekilde kapatılır

## Lisans

Bu proje MIT lisansı altında lisanslanmıştır. Detaylar için `LICENSE` dosyasına bakınız.

## İletişim

Sorularınız veya geri bildirimleriniz için:
- E-posta: [E-posta adresiniz]
- GitHub: [GitHub profiliniz]

## Katkıda Bulunma

1. Bu depoyu fork edin
2. Yeni bir özellik dalı oluşturun (`git checkout -b yeni-ozellik`)
3. Değişikliklerinizi commit edin (`git commit -am 'Yeni özellik eklendi'`)
4. Dalınıza push yapın (`git push origin yeni-ozellik`)
5. Bir Pull Request oluşturun
