# email-diagnostics

Bounce eden (teslim edilemeyen) e-posta satırlarını analiz edip her satıra **tek bir teşhis** yazan Python scripti. Önce e-posta adresinin kişi ve şirket adıyla tutarlılığını kontrol eder; tutarlıysa kişinin ilgili şirkette **Companies House** kayıtlarında aktif mi yoksa istifa etmiş mi olduğuna bakar.

Çıktı **tek bir Excel dosyasıdır**. Girdi dosyasına dokunulmaz.

---

## Önemli kısıt: hiçbir mail sunucusuna bağlanılmaz

Script SMTP handshake yapmaz, `RCPT TO` denemez, doğrulama servisi (ZeroBounce vb.) kullanmaz, DNS/MX sorgusu bile atmaz. Tek dış bağlantı Companies House REST API'sidir.

Bunun doğrudan sonucu: **bir adresin geçerli olduğu asla kanıtlanmaz.** `result` kolonu bir kanıt değil, "en olası açıklama"dır.

---

## Kurulum

```bash
pip install -r requirements.txt
```

İki paket yeterli: `openpyxl` ve `requests`.

pandas/numpy bilinçli olarak kullanılmadı — Python 3.6 / Windows ortamlarında kurulum sorunlu, ayrıca pandas `regnum` alanını sayıya çevirip baştaki sıfırları siliyor (`01234567` → `1234567`), bu da Companies House'tan 404 almanıza yol açıyor. Fuzzy matching kütüphanesi de yok; Damerau-Levenshtein saf Python ile yazıldı, derleyici gerektirmez.

Python 3.6.5 ve üzeri ile çalışır (`dataclasses`, f-string, walrus gibi 3.7+ özellikleri kullanılmadı).

## Companies House API anahtarı

Anahtar kodda **yazılı değildir**. Script sırayla şuralara bakar:

1. `CH_API_KEY` ortam değişkeni
2. Çalıştığınız klasördeki `.env` dosyası
3. Script'in yanındaki `.env` dosyası

| Ortam | Komut |
|---|---|
| Windows (kalıcı) | `setx CH_API_KEY "anahtar"` |
| Windows (geçici) | `set CH_API_KEY=anahtar` — tırnak koymayın |
| PowerShell | `$env:CH_API_KEY="anahtar"` |
| macOS / Linux | `export CH_API_KEY="anahtar"` |

> **Windows'ta en sık yaşanan sorun:** `setx` sadece **yeni açılan** işlemleri etkiler. VS Code kullanıyorsanız yeni bir terminal sekmesi açmak yetmez — VS Code ortamını başlatıldığı anda alır, **tamamen kapatıp yeniden açmanız** gerekir.

Uğraşmak istemiyorsanız alternatif: klasöre `.env` adlı bir dosya oluşturup içine tek satır yazın. Restart gerekmez, `.gitignore`'da olduğu için repoya da gitmez.

```
CH_API_KEY=buraya_anahtar
```

Anahtarı [Companies House Developer Hub](https://developer.company-information.service.gov.uk/) üzerinden bir "REST API" uygulaması oluşturarak alırsınız.

## Önce kurulumu doğrulayın

Yeni bir bilgisayarda ilk komut bu olsun:

```bash
python email_diagnostics.py check
```

Python sürümünü, paketleri ve API anahtarını sırayla kontrol eder; anahtar varsa Companies House'a **tek** test isteği atıp gerçekten çalıştığını doğrular. Eksik varsa nerede takıldığını ve nasıl düzelteceğinizi yazar.

Girdi dosyanızı da kontrol ettirebilirsiniz — okunabiliyor mu, zorunlu kolonlar var mı, kaç satır analiz edilecek:

```bash
python email_diagnostics.py check --input liste.csv
```

```
[OK]   Dosya okundu: 1240 satir
[OK]   Zorunlu kolonlarin hepsi var.
[OK]   Analiz edilecek satir: 183 / 1240
      Statu dagilimi:
        Delivered                      890  -
        Opened                         167  -
        Bounced                        142  analiz edilir
        Blocked                         41  -
        hard bounce                     41  analiz edilir
      Uyari: 3 satirda regnum bos.
```

`--skip-api` ile tamamen offline çalışır (test isteği bile atmaz).

## Çalıştırma

```bash
python email_diagnostics.py triage --input liste.csv --output liste_sonuc.csv --verbose
```

Windows'ta:

```bash
python email_diagnostics.py triage --input "C:\Users\adiniz\Downloads\liste.csv" --output "C:\Users\adiniz\Downloads\liste_sonuc.csv" --verbose
```

Girdi ve çıktı **`.xlsx` veya `.csv`** olabilir; uzantıya bakılarak otomatik seçilir. İkisi farklı da olabilir (CSV oku, Excel yaz).

### Bayraklar

| Bayrak | Ne yapar |
|---|---|
| `-i`, `--input` | Girdi dosyası (`.xlsx` / `.xlsm` / `.csv` / `.tsv`) |
| `-o`, `--output` | Çıktı dosyası — uzantısına göre CSV veya Excel yazılır |
| `-v`, `--verbose` | Her satırın kararını tek tek yazar |
| `--debug` | Çıktıya 11 denetim kolonu ekler |
| `--dry-run` | Companies House'a hiç gitmez — offline test |
| `--limit N` | Sadece ilk N problemli satırı işler — kota yakmadan deneme |
| `--sheet AD` | Excel sayfa adı (varsayılan: ilk sayfa) |
| `--delimiter` | CSV ayracı — verilmezse `;` `,` sekme `\|` arasından tahmin edilir |
| `--encoding` | CSV kodlaması — verilmezse `utf-8-sig`, `cp1254`, `cp1252` sırayla denenir |
| `--mode` | `typo_first` (varsayılan) veya `ch_first` |
| `--company-profile` | Resmî şirket adı + `company_dissolved` tespiti (şirket başına +1 istek) |
| `--workers N` | Paralel thread sayısı (varsayılan 4) |
| `--rate N` | Saniyedeki istek üst sınırı (varsayılan 1.8) |

Bayrak verilmezse dosyanın en üstündeki varsayılanlar kullanılır.

İlk deneme için önerilen komut — kotayı hiç harcamaz:

```bash
python email_diagnostics.py triage -i liste.csv -o deneme.csv --limit 50 --dry-run --debug --verbose
```

### Girdi dosyası

İlk satır başlık. Şu kolonlar zorunlu (büyük/küçük harf, boşluk ve BOM farkları tolere edilir):

`first_name` · `last_name` · `email` · `company` · `regnum` · `status`

`regnum` = Companies House şirket numarası. `status` = gönderim durumu.

CSV tarafında Türkiye/Avrupa Excel çıktıları da doğrudan çalışır: `;` ayracı ve `cp1254`/`cp1252` kodlaması otomatik algılanır, çıktı Excel'in Türkçe karakterleri doğru açması için `utf-8-sig` (BOM'lu) yazılır.

> **Not:** CSV çıktısını Excel'de açarsanız Excel `01234567` gibi `regnum` değerlerini yine sayıya çevirip baştaki sıfırı gizleyebilir. Bu Excel'in davranışıdır, dosyanın içeriği doğrudur — sorun yaşarsanız çıktıyı `.xlsx` olarak alın.

---

## Akış

1. **Yükleme + doğrulama** — zorunlu kolonlar eksikse hangilerinin eksik olduğunu söyleyen net bir hata verir.
2. **Statü filtresi** — sadece bounce ailesi satırlar işlenir. `blocked` bilinçli olarak **hariçtir**: spam filtresi/IP reputation kaynaklıdır, adresin yanlışlığı veya kişinin ayrılmasıyla ilgisi yoktur.
3. **Temizleme** — unvan (`Mr`, `Dr`), rol (`CEO`, `Director`), UK post-nominal (`MBE`, `FCA`) hem ad hem soyad alanından ayıklanır; parantezli lakaplar (`John (Jack)`) ayrı aday olarak saklanır; `last_name` tam ad içeriyorsa (`John Smith`) akıllıca bölünür; soyad ön ekleri (`van der`, `Mc`) korunur; şirket adından hukuki ekler atılır. **Orijinal değerler değiştirilmez**, temizlik yalnızca eşleştirme içindir.
4. **Typo kontrolü** — e-posta local part'ı beklenen kalıplarla karşılaştırılır. Typo bulunursa satır burada biter, API çağrılmaz.
5. **Companies House** — kalan satırlar için officer listesi sorgulanır, kişi **soyad öncelikli** eşleştirilir.
6. **Tek Excel çıktısı** + konsola sonuç dağılımı.

### Typo tespiti nasıl çalışıyor

Ad, soyad, orta ad, lakaplar ve baş harflerden `first.last`, `flast`, `f.last`, `lastfirst`, `f.m.last` gibi kalıplar 4 ayraçla (`.` `_` `-` ve bitişik) üretilir. Lakaplar iki yönlü bir sözlükle genişletilir (`Bob`↔`Robert`, `Jack`↔`John`, `Liz`↔`Elizabeth`…).

Karşılaştırma bilinçli olarak **muhafazakârdır** — zayıf farklara "typo" demez:

| Durum | Sonuç |
|---|---|
| Tam eşleşme | tutarlı |
| Sadece ayraç farkı (`johnsmith` ≡ `john.smith`) | tutarlı, typo **değil** |
| Lakap eşleşmesi (`jack.smith` ← John Smith) | tutarlı |
| 1–2 karakter mesafe (`jhon.smith`) | **typo** |
| Hiçbir kalıba yakın değil | `email_pattern_unrecognised` — typo **değil**, CH'ye devam |
| 5 karakterden kısa local part (`js@`) | asla typo değil; baş harf kombinasyonları denenir |

Mesafe hesabı **Damerau-Levenshtein**'dır: en yaygın yazım hatası iki harfin yer değiştirmesidir (`jhon` ← `john`), düz Levenshtein bunu 2 hata sayıp eşiği kaçırır.

Domain kontrolünde sırasıyla tam eşleşme → en az 2 karakterlik anlamlı ek ile içerme (`acme` → `acmegroup` ✔) → 1–2 mesafe typo (`acmee` ✘) bakılır. Şirket adından kısaltma da üretilir: *Ali Veli Zeynep Ltd* → `avz`, `alivelizeynep`, `aliveli`. Alakasız domain **typo sayılmaz** — holding/grup domaini olabilir.

### Officer eşleştirme

Soyad çapadır: önce soyad tutmalı (tam eşleşme veya ≤1 mesafe), sonra ad değerlendirilir. Companies House'un yapısal `name_elements` alanı kullanılır, string bölme yapılmaz. `former_names` de taranır — evlilik nedeniyle değişmiş soyadları böyle yakalanır.

- Soyad + ad tuttu → **confident** (`active_officer_match` / `resigned_officer_match`)
- Sadece baş harf veya sadece soyad tuttu → **possible**
- Aynı soyadlı birden fazla officer (aile şirketleri) → belirsiz olarak işaretlenir, güven düşürülür
- Aynı kişinin hem istifa hem aktif kaydı varsa **aktif kazanır**
- Kurumsal officer'lar (`corporate-director`) atlanır

---

## Çıktı kolonları

Orijinal kolonların tamamı korunur, sonlarına eklenir:

| Kolon | İçerik |
|---|---|
| `result` | Birincil teşhis |
| `result_reason` | Kısa destekleyici gerekçe |
| `ch_officer_name` | Companies House'tan gelen resmî tam ad (orta adlar dahil) |
| `ch_officer_status` | `active` / `resigned` / `possible_active` / `possible_resigned` / `not_found` / `lookup_failed` / `not_checked` |

`DEBUG = True` yapılırsa 11 denetim kolonu daha eklenir (temizlenmiş ad/soyad, lakaplar, şirket token'ları, eşleşen kalıp, mesafe, kullanılan regnum…).

### `result` değerleri

`missing_email` · `malformed_email` · `first_name_typo` · `surname_typo` · `first_name_and_surname_typo` · `domain_typo` · `active_officer_match: <isim>` · `resigned_officer_match: <isim>` · `possible_officer_match_active: <isim>` · `possible_officer_match_resigned: <isim>` · `no_officer_match_found` · `company_not_found` · `company_dissolved` · `companies_house_lookup_failed` · `missing_regnum`

### `result_reason` değerleri

`email_matches_expected_pattern` · `generic_mailbox` · `personal_email_domain` · `email_pattern_unrecognised` · `domain_not_matched` · `close_to_expected_pattern` · `matched_including_middle_name` · `surname_only_match` · `multiple_possible_officers` · `api_error` · `name_fields_empty`

---

## Ayarlar

Hepsi dosyanın en üstünde:

Çoğu ayarın komut satırı karşılığı var (yukarıdaki tabloya bakın). Yalnızca dosyadan değiştirilebilenler:

| Ayar | Ne işe yarar |
|---|---|
| `PROBLEMATIC_STATUS_KEYWORDS` | Hangi statülerin analiz edileceği |
| `EXCLUDED_STATUS_KEYWORDS` | Hangi statülerin dışlanacağı (`blocked` burada) |
| `NICKNAME_GROUPS` | Lakap sözlüğü — istediğiniz kadar genişletin |
| `TYPO_MAX_DISTANCE_*` | Typo eşikleri |
| `GENERIC_MAILBOXES` | `info@`, `accounts@` gibi kişiye ait olmayan kutular |
| `COMPANY_STOPWORDS` | Şirket adından atılacak hukuki ekler |
| `FREE_EMAIL_DOMAINS` | Kişisel e-posta sağlayıcıları |

## Rate limit ve performans

Companies House limiti **5 dakikada 600 istek** (= 2/sn). Script 4 thread ve saniyede 1.8 istek throttle ile çalışır; thread'ler tavanı *doldurmaya* yarar, tavanı aşamaz. Aynı `regnum` bir kez sorgulanır. `401` alınırsa tüm çalışma anında durur (binlerce satıra boşuna `lookup_failed` yazmamak için), `429` ve `5xx` için exponential backoff uygulanır, officer listesi sayfalanır (35'ten fazla officer'ı olan şirketler için şart).

Kaba tahmin: 1.000 farklı şirket ≈ 10 dakika.

## Testler

```bash
python test_logic.py   # birim testler: temizleme, typo, officer eşleştirme, regnum, statü
python test_e2e.py     # uçtan uca: gerçek .xlsx oluşturur, çalıştırır, çıktıyı doğrular
```

Ağ bağlantısı gerektirmezler.

---

## Bilinen sınırlar

- `result` bir kanıt değil, en olası açıklamadır. Bounce'un sebebi ne typo ne istifa olabilir (kutu dolu, kutu silinmiş, spam filtresi, sunucu arızası).
- Domain eşleştirmesi şirketin **tescilli** adına dayanır. Marka/ticari ad farklıysa `domain_not_matched` çıkar — bu bilinçli olarak typo sayılmaz.
- Companies House yalnızca **officer** (direktör, sekreter, LLP üyesi) kayıtlarını tutar. Listeniz officer'lardan oluşmuyorsa `no_officer_match_found` baskın çıkar ve anlam taşımaz.
- Çıktı dosyası kişisel veri içerir. GDPR kapsamında saklama ve paylaşım sorumluluğu sizdedir.

## Lisans

MIT
