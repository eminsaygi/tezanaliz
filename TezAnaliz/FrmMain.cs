using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace TezAnaliz
{
    public partial class FrmMain : DevExpress.XtraEditors.XtraForm
    {
        public FrmMain()
        {
            InitializeComponent();
            Thread.Sleep(3000);
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            btnLoad.Enabled = false;//Yükle butonu kapatılıyor .
            btnTest.Enabled = false;//Başla butonu kapatılıyor .
            if (thesisProcess.IsBusy)
            {
                MessageBox.Show("İşlem devam ediyor. Lütfen Bekleyiniz");
                return;
            }
            listBoxResult.Items.Clear();//Tez incelendiyse sonuç listesi temizleniyor.
            listBoxResult.Items.Add("--------------------------------------------------");
            listBoxResult.Items.Add("| Tez Analizi süreci başladı. Lütfen bekle! |");//listeye  ilk mesajı yazdırdık.
            listBoxResult.Items.Add("--------------------------------------------------");
            thesisProcess.RunWorkerAsync();//Bir thread başlatıyoruz.
        }

        private void thesisRequest(object sender, ProgressChangedEventArgs e)
        {
            //İşlemin sürecini texte yüzdelik olarak yazdırıyoruz.
            labelControl1.Text = "%" + e.ProgressPercentage;
        }

        private void pictureEdit1_Click(object sender, EventArgs e)
        {
            this.Close(); //Simgeye basınca formu kapatıyoruz.
        }

        private void thesisStop(object sender, RunWorkerCompletedEventArgs e)
        {
            listBoxResult.Items.Add("--------------------------------------------------");
            listBoxResult.Items.Add("|  Tez Analiz süreci bitti!  |");//listeye  son mesajı yazdırdık.
            listBoxResult.Items.Add("--------------------------------------------------");
            labelControl1.Text = "%100";
            this.Text = "Fırat Üniversitesi Tez Analiz Aracı";
            btnTest.Enabled = true;//başla butonu açılıyor .
            btnLoad.Enabled = true;//Yükle butonu açılıyor .
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileExplorer = new OpenFileDialog();
            fileExplorer.Filter = "World Dosyası |*.doc| World Dosyası|*.docx";//Uzantı .doc yada docx olarak ayarlanmıştır.
            fileExplorer.FilterIndex = 2;
            if (fileExplorer.ShowDialog() == DialogResult.OK)//Tamam ise;
            {
                txtPath.Text = fileExplorer.FileName;// Seçilen dosya atanır.
            }
        }

        private void thesisAnimation(object sender, DoWorkEventArgs e)//Thread basladığında fonksiyon çalışır .
        {
            int hata = 0;

            thesisProcess.ReportProgress(1);//Yüzdeyi göstermek için kullandığımız method gönderir.

            Microsoft.Office.Interop.Word.Application app = new Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(txtPath.Text);//Document objesinin içine yükleme işlemi.
            Word.WdStatistic stati = Word.WdStatistic.wdStatisticPages;
            object empty = System.Reflection.Missing.Value;
            int countN = doc.ComputeStatistics(stati, ref empty);//Sayfa sayısını countN değişkenine atıyoruz.

            thesisProcess.ReportProgress(3);

            //Sayfa sayısını kontrol ettik.
            if (countN < 40 || countN > 180)
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");
                    this.listBoxResult.Items.Add("Tez sayfa uzunluğu 40 - 180 sayfa arasında olmalıdır! Mevcut sayfa sayısı:" + countN);
                }));
            }

            float solKenar = doc.PageSetup.LeftMargin;//Sol boşluk,
            float sagKenar = doc.PageSetup.RightMargin;//Sağ boşluk,
            float ustKenar = doc.PageSetup.TopMargin;//Üst boşluk,
            float altKenar = doc.PageSetup.BottomMargin;//Alt boşluk değerleri dokümandan okunuyor.

            if (sagKenar != 70.9f)//Sağ boşluk kontrolü.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");
                    this.listBoxResult.Items.Add("Sayfa Sağ kenar boşluğu 2.5 cm olmalıdır!");
                }));
            }

            if (altKenar != 70.9f)//Alt Kenar boşluk kontrolü.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");
                    this.listBoxResult.Items.Add("Sayfa Alt kenar boşluğu 2.5 cm olmalıdır!");
                }));
            }

            if (solKenar != 92.15f)//Sol Kenar boşluk kontrolü.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");
                    this.listBoxResult.Items.Add("Sayfa Sol kenar boşluğu 3.25 cm olmalıdır!");
                }));
            }

            if (ustKenar != 85.05f)//Üst Kenar boşluk kontrolü.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");
                    this.listBoxResult.Items.Add("Sayfa Üst kenar boşluğu 3.0 cm olmalıdır!");
                }));
            }
            thesisProcess.ReportProgress(5);

            int paragrafSayisi = doc.Paragraphs.Count;

            int sayi = 0;
            int timeFontCalc = 0;
            int fontBigCalc = 0;

            bool onsozC = false;
            bool contentsControl = false;
            bool sumControl = false;
            bool refControl = false;
            bool statControl = false;

            bool abstractControl = false;
            bool sekilListControl = false;
            bool attachListControl = false;
            bool symbolControl = false;

            int sekilSayisi = doc.InlineShapes.Count;
            int detaySekilSayisi = doc.Shapes.Count;

            int miktarTablo = 0;

            int mSekil = 0;

            Hashtable circle = new Hashtable();
            Hashtable crcTable = new Hashtable();

            double sekilSay = 0;

            foreach (Paragraph paragrafSinifi in doc.Paragraphs)//Paragrafların her biri tek tek okunup paragrafSinifi objesinin içine atılıyor.
            {
                if (paragrafSinifi.Range.Font.Name == "Times New Roman")//Eğer font Times New Roman ise;
                    timeFontCalc++;//Times New Roman sayısı  arttırılıyor.

                if (paragrafSinifi.Range.Font.Size == 11)//Font size'ı 11 ise;
                    fontBigCalc++;//Font size'ı kontrol eden değiken  arttırılıyor.
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.Contents)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Contents);
                    object basla2 = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Contents) + 1;

                    object bitis = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Contents) + Sabitler.Contents.Length;

                    Word.Range cevreFont = doc.Range(ref basla, ref basla2);
                    Word.Range digerCevre = doc.Range(ref basla2, ref bitis);

                    float textS = cevreFont.Font.Size;
                    float textO = digerCevre.Font.Size;

                    if (textS != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalıdır! Bölüm :'İÇİNDEKİLER'");//Listeye Eklenir.
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'İÇİNDEKİLER'");//Listeye Eklenir
                        }));
                    }

                    contentsControl = true;
                }
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.Summary)//Özet textini çağırır
                {
                    object başla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Summary);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Summary) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Summary) + Sabitler.Summary.Length;

                    Word.Range ilkKar = doc.Range(ref başla, ref baslaR);
                    Word.Range karD = doc.Range(ref baslaR, ref son);

                    float textS = ilkKar.Font.Size;
                    float textO = karD.Font.Size;

                    if (textS != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'Özet'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'Özet'");//Listeye Eklenir
                        }));
                    }

                    sumControl = true;
                }
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.Preface)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Preface);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Preface) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Preface) + Sabitler.Preface.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreA = doc.Range(ref baslaR, ref son);

                    float textS = cevreR.Font.Size;
                    float textO = cevreA.Font.Size;

                    if (textS != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'ÖNSÖZ'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'ÖNSÖZ'");//Listeye Eklenir
                        }));
                    }
                    onsozC = true;
                }
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.Resources)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Resources);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Resources) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Resources) + Sabitler.Resources.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreO = doc.Range(ref baslaR, ref son);

                    float textF = cevreR.Font.Size;
                    float textO = cevreO.Font.Size;

                    if (textF != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'Resources'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");
                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'Resources'");//Listeye Eklenir
                        }));
                    }
                    refControl = true;
                }
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.ABSTRACT)
                {
                    object start = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.ABSTRACT);
                    object startplusone = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.ABSTRACT) + 1;

                    object end = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.ABSTRACT) + Sabitler.ABSTRACT.Length;

                    Word.Range rangeFirstChar = doc.Range(ref start, ref startplusone);
                    Word.Range rangeothers = doc.Range(ref startplusone, ref end);

                    float textsizefirst = rangeFirstChar.Font.Size;
                    float textsizeothers = rangeothers.Font.Size;

                    if (textsizefirst != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'ABSTRACT'");//Listeye Eklenir
                        }));
                    }

                    if (textsizeothers != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'ABSTRACT'");//Listeye Eklenir
                        }));
                    }
                    abstractControl = true;
                }

                if (paragrafSinifi.Range.Text.Trim() == Sabitler.Declaration)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Declaration);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Declaration) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Declaration) + Sabitler.Declaration.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreO = doc.Range(ref baslaR, ref son);

                    float textR = cevreR.Font.Size;
                    float textO = cevreO.Font.Size;
                    if (textR != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'Declaration'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'Declaration'");//Listeye Eklenir
                        }));
                    }
                    statControl = true;
                }
                if (paragrafSinifi.Range.Text.Trim() == Sabitler.FiguresList)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.FiguresList);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.FiguresList) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.FiguresList) + Sabitler.FiguresList.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreO = doc.Range(ref baslaR, ref son);

                    float textR = cevreR.Font.Size;
                    float textO = cevreO.Font.Size;
                    if (textR != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'ŞEKİLLER LİSTESİ'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'ŞEKİLLER LİSTESİ'");//Listeye Eklenir
                        }));
                    }
                    sekilListControl = true;
                }

                if (paragrafSinifi.Range.Text.Trim() == Sabitler.AttachList)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.AttachList);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.AttachList) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.AttachList) + Sabitler.AttachList.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreO = doc.Range(ref baslaR, ref son);

                    float textR = cevreR.Font.Size;
                    float textO = cevreO.Font.Size;

                    if (textR != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'EKLER LİSTESİ'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'EKLER LİSTESİ'");//Listeye Eklenir.
                        }));
                    }
                    attachListControl = true;
                }

                if (paragrafSinifi.Range.Text.Trim() == Sabitler.SymbolsAbrevv)
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.SymbolsAbrevv);
                    object baslaR = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.SymbolsAbrevv) + 1;

                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.SymbolsAbrevv) + Sabitler.SymbolsAbrevv.Length;

                    Word.Range cevreR = doc.Range(ref basla, ref baslaR);
                    Word.Range cevreO = doc.Range(ref baslaR, ref son);

                    float textR = cevreR.Font.Size;
                    float textO = cevreO.Font.Size;

                    if (textR != 16)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("İlk harf 16 punto olmalı! Bölüm :'SİMGELER VE KISALTMALAR'");//Listeye Eklenir
                        }));
                    }

                    if (textO != 13)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Başlığın ilk harfinden sonra gelen harfler 13 punto olmalı! Bölüm :'SİMGELER VE KISALTMALAR'");//Listeye Eklenir
                        }));
                    }
                    symbolControl = true;
                }

                if (paragrafSinifi.Range.Text.Contains(Sabitler.Figure))
                {
                    //Şekil 2.3.
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Figure);
                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Figure) + Sabitler.Figure.Length + 5;
                    Word.Range cevreSekil = doc.Range(ref basla, ref son);

                    bool sekilI = Yardimcilar.TextIsSekil(cevreSekil.Text);

                    if (sekilI && !circle.ContainsKey(cevreSekil.Text))//Şekil  kontrolü yapılıyor.
                    {
                        mSekil++;//Şekil Sayacı arttırılıyor.
                        circle[cevreSekil] = cevreSekil.Text;//Şekil atılıyor.

                        double sekilN = Double.Parse(cevreSekil.Text.Substring(6, 3));// 4.3

                        if (sekilN < sekilSay)//sekilSay en son şeklin numarasını tutar.Son şekilden küçükse hata vardır.
                        {
                            hata++;
                            this.Invoke(new MethodInvoker(() =>
                            {
                                listBoxResult.Items.Add("--------------------------------------------------");

                                this.listBoxResult.Items.Add("Şekil numaraları sırayla olmalıdır!");//Listeye Eklenir
                            }));
                        }
                        else sekilSay = sekilN;
                    }

                    if (sekilI && cevreSekil.Font.Size < 10)//Eğer şekil gelmiş ve font size'ı 10'dan küçükse Listeye Eklenir
                    {
                        hata++;

                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Şekil kısımları en az 10 boyutunda olmalı!");
                        }));
                    }
                }
                //Declaration
                if (paragrafSinifi.Range.Text.Contains(Sabitler.Table))//Bloğun tamamı tablo için de yapılıyor.
                {
                    object basla = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Table);
                    object son = paragrafSinifi.Range.Start + paragrafSinifi.Range.Text.IndexOf(Sabitler.Table) + Sabitler.Table.Length + 5;
                    Word.Range cevreTablo = doc.Range(ref basla, ref son);

                    bool tabloI = Yardimcilar.TextIsTablo(cevreTablo.Text);
                    if (tabloI && !crcTable.ContainsKey(cevreTablo.Text))
                    {
                        miktarTablo++;
                        crcTable[cevreTablo] = cevreTablo.Text;
                    }

                    if (tabloI && cevreTablo.Font.Size < 10)
                    {
                        hata++;
                        this.Invoke(new MethodInvoker(() =>
                        {
                            listBoxResult.Items.Add("--------------------------------------------------");

                            this.listBoxResult.Items.Add("Tablo kısımları en az 10 boyutunda olmalı!");
                        }));
                    }
                }

                thesisProcess.ReportProgress(5 + (95 * sayi / doc.Paragraphs.Count));//Yüzde hesaplama Toplam paragraf indexine oranını buluyoruz.

                sayi++;
            }

            if ((sekilSayisi + detaySekilSayisi) < (miktarTablo + mSekil))//Tablo ve şekil sayı'ından küçükse o zaman tüm tablo ve şekillere isim verilmemiştir.
            {//Bu durumda ekrana hata bastırılır.
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Tüm şekil ve tabloların altına Şekil ve Tablo numarası yazılmalıdır!");
                }));
            }

            if (!statControl)//Yukarıdaki gibi
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Declaration bölümü eksik!");
                }));
            }
            if (!onsozC)//Önsöz değişkenine yukarıda atama yapıldı.
            {
                hata++;//hata sayısı bir arttırılır.

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("ÖNSÖZ bölümü eksik!");// Önsöz bölümü eksik Listeye Eklenir
                }));
            }
            if (!contentsControl)//Eğer içindekiler değişkenine true atanmadıysa;
            {//Listeye Eklenir
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("İÇİNDEKİLER bölümü eksik!");
                }));
            }
            if (!sumControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("ÖZET bölümü eksik!");
                }));
            }
            if (!abstractControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("ABSTRACT bölümü eksik!");
                }));
            }

            if (!sekilListControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("ŞEKİLLER LİSTESİ bölümü eksik!");
                }));
            }
            if (!attachListControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("EKLER LİSTESİ bölümü eksik!");
                }));
            }

            if (!symbolControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("SIMGELER VE KISALTMALAR bölümü eksik!");
                }));
            }
            if (!refControl)//Aynı işlemler gerçekleştiriliyor
            {
                hata++;
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Resources bölümü eksik!");
                }));
            }
            if (timeFontCalc * 2 < paragrafSayisi)//Times New Roman ile yazılmamışsa ekrana hata bastırılır.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Tezin yazım fontu Times New Roman olmalıdır! Tezdeki Times New Roman oranı = %" + 100 * timeFontCalc / paragrafSayisi);
                }));
            }

            if (fontBigCalc * 1.5 < paragrafSayisi)//Tez içeriğindeki yazıların boyutu 11 punto olmalıdır.
            {
                hata++;

                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Tezin font büyüklüğü 11 punto olmalıdır! Tezde bulunan 11 punto oranı= %" + 100 * fontBigCalc / paragrafSayisi);
                }));
            }

            if (hata == 0)//Eğer hata sayısı 0 ise tezinizde hata bulunamadı diye mesaj gösteriyoruz.
            {
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add("Tezinizde hiçbir hata bulunamadı!");
                }));
            }
            else
            {
                this.Invoke(new MethodInvoker(() =>
                {
                    listBoxResult.Items.Add("--------------------------------------------------");

                    this.listBoxResult.Items.Add(string.Format("Tezinizde bulunan hata başlığı sayısı:{0}", hata));//Hata sayısı 0'dan farklıysa da tezde bulunan toplam hata sayısı gösteriliyor.
                }));
            }

            doc.Close();
            doc = null;
            //return num;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            frmHakkimda frmHakkimda = new frmHakkimda();
            frmHakkimda.Show();
        }
    }
}