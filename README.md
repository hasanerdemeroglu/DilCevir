# DILCEVIR – Excel Translation Add-in

Excel hücrelerindeki metinleri farklı dillere çevirmek için geliştirilmiş basit bir VBA fonksiyonudur. Hücre formülü olarak kullanılır ve formül aşağı doğru kopyalanarak bir sütundaki tüm metinler çevrilebilir.

## Özellikler

- Hücre formülü olarak kullanım
- Otomatik kaynak dil algılama
- Birden fazla hedef dil desteği
- Hücre referanslarıyla çalışma
- Formülü aşağı doğru sürükleyerek toplu çeviri
- Türkçe karakter desteği
- UTF-8 desteği
- İnternet bağlantısı üzerinden çeviri
- Google Translate altyapısını kullanır

## Kullanım

Excel hücresine aşağıdaki formülü yazın:

```excel
=DILCEVIR(A1;"EN")
```

Bu formül A1 hücresindeki metni İngilizceye çevirir.

### Örnekler

| Formül | Hedef Dil |
|---|---|
| `=DILCEVIR(A1;"EN")` | İngilizce |
| `=DILCEVIR(A1;"AZ")` | Azerbaycanca |
| `=DILCEVIR(A1;"RU")` | Rusça |
| `=DILCEVIR(A1;"DE")` | Almanca |
| `=DILCEVIR(A1;"FR")` | Fransızca |
| `=DILCEVIR(A1;"AR")` | Arapça |
| `=DILCEVIR(A1;"ES")` | İspanyolca |

## Toplu Çeviri

Örneğin A sütununda kaynak metinler bulunuyorsa:

| A | B |
|---|---|
| Merhaba | `=DILCEVIR(A1;"EN")` |
| Günaydın | `=DILCEVIR(A2;"EN")` |
| İyi akşamlar | `=DILCEVIR(A3;"EN")` |

B1 hücresine formülü yazdıktan sonra aşağı doğru sürükleyerek diğer satırlara uygulayabilirsiniz.

## Kurulum

1. Excel'i açın.
2. `Alt + F11` tuşlarına basarak VBA editörünü açın.
3. `Insert > Module` seçeneğiyle yeni modül oluşturun.
4. DILCEVIR VBA kodunu modüle yapıştırın.
5. Dosyayı Excel Add-In formatında kaydedin (`.xlam`).
6. Excel'de `Dosya > Seçenekler > Eklentiler` bölümüne gidin.
7. Alt kısımdaki `Yönet: Excel Eklentileri` seçeneğinden `Git` butonuna basın.
8. Eklentinizi seçip etkinleştirin.

## Gereksinimler

- Microsoft Excel
- Windows
- İnternet bağlantısı
- VBA desteği

## Notlar

- Çeviri işlemi internet bağlantısı gerektirir.
- Google Translate'in resmi olmayan ücretsiz endpoint'i kullanılmaktadır.
- Çok sayıda hücreyi aynı anda çevirmek, servis tarafından hız sınırlandırmasına neden olabilir.
- Formüller Excel tarafından yeniden hesaplandığında çeviri isteği tekrar gönderilebilir.
- Çeviri sonuçları hedef hücrelerde formül olarak kalır.

## Dil Kodları

Google Translate tarafından kullanılan dil kodları kullanılmaktadır.

Örnek:

- `TR` – Türkçe
- `EN` – İngilizce
- `AZ` – Azerbaycanca
- `RU` – Rusça
- `DE` – Almanca
- `FR` – Fransızca
- `ES` – İspanyolca
- `IT` – İtalyanca
- `AR` – Arapça
- `ZH` – Çince
- `JA` – Japonca
- `KO` – Korece

## Lisans

Bu proje kişisel kullanım ve geliştirme amacıyla hazırlanmıştır.

---

**DILCEVIR**  
Excel içinde pratik ve hızlı çeviri.
