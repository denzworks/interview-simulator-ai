# Mülakat Simülatörü AI – Yapay Zeka Destekli Mülakat Hazırlık Sistemi

İş ilanını analiz eden, sana özel sorular üreten, cevaplarını anlık değerlendiren ve detaylı işe alım raporu oluşturan bir web uygulaması.

---

## Ne Yapar?

- Yapıştırdığın veya yüklediğin iş ilanını AI ile analiz eder; pozisyon seviyesini, zorunlu ve tercih edilen becerileri otomatik çıkarır
- Analiz sonucuna göre 10 teknik, 5 davranışsal ve 5 problem çözme sorusu üretir
- Her cevabı anlık olarak teknik doğruluk, açıklık, eksiksizlik ve iletişim boyutlarında 1–10 üzerinden puanlar
- Mülakat bitince genel, teknik, davranışsal, problem çözme ve iletişim skorlarını hesaplar
- Güçlü yönleri, zayıflıkları, eksik becerileri ve geliştirme planını raporlar
- Nihai işe alım tavsiyesi üretir: Kesinlikle İşe Al / İşe Al / Sınırda / İşe Alma
- Raporu PDF ve DOCX formatında dışa aktarır

---

## Proje Yapısı

```
root/
├── interview_simulator.py   → Tüm uygulama (tek dosya)
└── requirements.txt         → Gerekli Python kütüphaneleri
```

---

## Dosya Ne İş Yapar?

### `interview_simulator.py`

Tüm uygulama tek bir dosyada toplanmıştır. İçerdiği fonksiyonlar:

| Fonksiyon | Açıklama |
|-----------|----------|
| `call_openrouter()` | OpenRouter API'ye istek gönderir, hata durumunda kullanıcıyı bilgilendirir |
| `extract_json_from_text()` | AI yanıtından JSON bloğunu güvenli şekilde ayrıştırır |
| `analyze_job_description()` | İş ilanından pozisyon, seviye, beceriler ve deneyim yılını çıkarır |
| `generate_interview_questions()` | Pozisyona özel 20 soru (teknik/davranışsal/problem çözme) üretir |
| `evaluate_answer()` | Verilen cevabı 4 boyutta puanlar ve geri bildirim üretir |
| `calculate_final_scores()` | Tüm cevap puanlarından kategori bazlı nihai skorları hesaplar |
| `generate_hiring_recommendation()` | Genel skora göre işe alım tavsiyesi belirler |
| `generate_interview_report()` | Güçlü/zayıf yönler, eksik beceriler ve gelişim planını AI ile oluşturur |
| `export_docx_report()` | python-docx ile stillendirilmiş Word raporu üretir |
| `export_pdf_report()` | ReportLab ile renkli ve yapılandırılmış PDF raporu üretir |
| `render_home_page()` | İş ilanı giriş ekranını ve analiz/başlatma akışını yönetir |
| `render_interview_page()` | Soru–cevap akışını, ilerleme çubuğunu ve anlık puanlamayı yönetir |
| `render_analysis_page()` | Radar chart, dağılım grafiği ve Q&A inceleme panelini render eder |
| `render_report_page()` | İşe alım tavsiyesini, özet raporu ve indirme butonlarını gösterir |
| `render_sidebar()` | API anahtarı, model seçimi, navigasyon ve hızlı puan özeti |
| `main()` | Session state yönetimi ve sayfa yönlendirmesi |

---

## Mülakat Akışı

```
İş ilanı gir / yükle
        ↓
AI ilanı analiz eder (pozisyon, seviye, beceriler)
        ↓
AI 20 soru üretir
        ↓
Mülakat başlar — her cevap anlık puanlanır
        ↓
Tüm sorular tamamlanır
        ↓
AI nihai rapor + işe alım tavsiyesi üretir
        ↓
PDF veya DOCX olarak indir
```

---

## Skor Hesaplama

Her cevap 4 boyutta 1–10 üzerinden değerlendirilir:

| Boyut | Açıklama |
|-------|----------|
| Teknik Doğruluk | Cevabın teknik içerik kalitesi |
| Açıklık | Cevabın anlaşılır ve düzenli olması |
| Eksiksizlik | Soruyu tüm yönleriyle karşılama oranı |
| İletişim | Kendini ifade etme ve sunum kalitesi |

Nihai skorlar 100 üzerinden hesaplanır ve renk kodlanır:

- 🟢 **80 ve üzeri** → Mükemmel
- 🟡 **60–79** → İyi
- 🔴 **0–59** → Geliştirilmeli

---

## İşe Alım Tavsiyesi

| Tavsiye | Puan Aralığı |
|---------|-------------|
| 🟢 Kesinlikle İşe Al | 80+ |
| 🔵 İşe Al | 65–79 |
| 🟡 Sınırda | 45–64 |
| 🔴 İşe Alma | 0–44 |

---

## Kurulum

**1. Repoyu klonla:**
```bash
git clone https://github.com/kullanici/interview-simulator-ai.git
cd interview-simulator-ai
```

**2. Kütüphaneleri yükle:**
```bash
pip install -r requirements.txt
```

**3. API anahtarını ayarla:**

[openrouter.ai/keys](https://openrouter.ai/keys) adresine ücretsiz kayıt ol ve API anahtarını kopyala.

Ortam değişkeni olarak tanımla (önerilen):
```bash
# Windows
set OPENROUTER_API_KEY=sk-or-v1-...

# Linux / macOS
export OPENROUTER_API_KEY=sk-or-v1-...
```

Alternatif olarak uygulamayı başlattıktan sonra sol panelden de girebilirsin.

**4. Uygulamayı başlat:**
```bash
streamlit run interview_simulator.py
```

---

## Kullanım

1. Sol panelden OpenRouter API anahtarını gir (ortam değişkeni tanımladıysan otomatik gelir)
2. **Ana Sayfa**'da iş ilanını metin alanına yapıştır veya TXT/DOCX/PDF olarak yükle
3. **İş İlanını Analiz Et** butonuna bas
4. Analiz tamamlanınca **Mülakatı Başlat** butonuna bas
5. Her soruyu dikkatlice oku, cevabını yaz ve **Gönder & Sonraki** butonuna bas
6. Tüm sorular bitince **Analiz** sayfasında performansını incele
7. **Rapor** sayfasından PDF veya DOCX olarak indir

---

## Sayfalar

| Sayfa | İçerik |
|-------|--------|
| 🏠 Ana Sayfa | İş ilanı giriş alanı, dosya yükleme, analiz ve başlatma |
| 🎤 Mülakat | 20 soru, anlık puanlama, kategori etiketi, ilerleme çubuğu |
| 📊 Analiz | Puan kartları, radar chart, dağılım grafiği, Q&A inceleme |
| 📋 Rapor | İşe alım tavsiyesi, özet, gelişim planı, PDF/DOCX indirme |

---

## Gereksinimler

- Python 3.10+
- İnternet bağlantısı (API çağrısı için)
- OpenRouter hesabı (ücretsiz)

### Kullanılan Kütüphaneler

```
streamlit>=1.37.0
requests>=2.32.3
pandas>=2.2.3
numpy>=2.0.0
plotly>=5.22.0
python-docx>=1.1.2
reportlab>=4.2.2
pdfplumber>=0.11.0
```

---

## Kullanılan Model

Varsayılan model: `openrouter/auto:free`

Sol panelden aşağıdaki modeller arasında seçim yapılabilir:

- `openrouter/auto:free` *(varsayılan, ücretsiz)*
- `mistralai/mistral-7b-instruct:free`
- `google/gemma-2-9b-it:free`
- `meta-llama/llama-3.1-8b-instruct:free`
- `anthropic/claude-3-haiku`
- `openai/gpt-4o-mini`

---

## Notlar

- Uygulama metin tabanlı PDF dosyalarını destekler. Görüntü tabanlı (taranmış) PDF'lerde metin çıkarılamaz; bu durumda DOCX veya TXT formatında yükleme önerilir.
- Ücretsiz modeller rate limit'e sahiptir. Her cevap değerlendirmesi ayrı bir API çağrısı yaptığından yoğun kullanımda kısa gecikmeler yaşanabilir.
- API anahtarını ortam değişkeni olarak tanımlamak, her oturumda yeniden girmek zorunda kalmamayı sağlar.
- Mülakat tamamlandıktan sonra sol paneldeki **Mülakatı Sıfırla** butonu ile yeni bir oturum başlatılabilir; API anahtarı ve model seçimi korunur.
