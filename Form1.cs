using System.Drawing;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;

using gfoidl.Imaging;
using System;

namespace PhotoGrouper
{
    public partial class Form1 : Form

    {


        int currentcell = 0;
        int nextcell = 0;
        string yol = "C:/Users/UGUR/Desktop/bilgisayar organizasyonu";
        string DosyaYolu;
        string DosyaAdi;
        DataTable dt;
        //
        private Image _originalImage;
        private bool _selecting;
        private Rectangle _selection;

        public DataTable ToDataTable(ExcelApp.Range range, int rows, int cols)
        {
            DataTable table = new DataTable();
            for (int i = 1; i <= rows; i++)
            {
                if (i == 1)
                { // ilk satırı Sutun Adları olarak kullanıldığından
                  // bunları Sutün Adları Olarak Kaydediyoruz.
                    for (int j = 1; j <= cols; j++)
                    {
                        //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            table.Columns.Add(range.Cells[i, j].Value2.ToString());
                        else //Boş olduğunda Kaçınsı Sutünsa Adı veriliyor.
                            table.Columns.Add(j.ToString() + ".Sütun");
                    }
                    continue;
                }
                //Yukarıda Sütunlar eklendi
                // onun şemasına göre yeni bir satır oluşturuyoruz. 
                //Okunan verileri yan yana sıralamak için
                var yeniSatir = table.NewRow();
                for (int j = 1; j <= cols; j++)
                {
                    //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        yeniSatir[j - 1] = range.Cells[i, j].Value2.ToString();
                    else // İçeriği boş hücrede hata vermesini önlemek için
                        yeniSatir[j - 1] = String.Empty;
                }
                table.Rows.Add(yeniSatir);
            }
            return table;
        }

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            _originalImage = pictureBox1.Image.Clone() as Image;

        }
       
       
        //---------------------------------------------------------------------

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;// seçilen dosyanın tüm yolunu verir
                DosyaAdi = file.SafeFileName;// seçilen dosyanın adını verir.
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                { //Excel Yüklümü Kontrolü Yapılmaktadır.
                    MessageBox.Show("Excel yüklü değil.");
                    return;
                }
                //Excel Dosyası Açılıyor.
                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(DosyaYolu);
                //Excel Dosyasının Sayfası Seçilir.
                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                //Excel Dosyasının ne kadar satır ve sütun kaplıyorsa tüm alanları alır.
                ExcelApp.Range excelRange = excelSheet.UsedRange;
                int satirSayisi = excelRange.Rows.Count; //Sayfanın satır sayısını alır.
                int sutunSayisi = excelRange.Columns.Count;//Sayfanın sütun sayısını alır.
                dt = ToDataTable(excelRange, satirSayisi, sutunSayisi);
                dataGridView1.DataSource = dt;
                dataGridView1.Refresh();
                //Okuduktan Sonra Excel Uygulamasını Kapatıyoruz.
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            else
            {
                MessageBox.Show("Dosya Seçilemedi.");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                foreach (string oge in openFileDialog2.FileNames)
                {
                    listView1.Items.Add(System.IO.Path.GetFileName(oge));
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (var item in dataGridView1.Rows)
            {
                File.Copy(listView1.Items[0].ToString(), "C:/Users/UGUR/Desktop/asdsadadas" + dataGridView1.CurrentRow.Selected);
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = _originalImage.Clone() as Image;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (pictureBox1.Image != null)
                {
                    Image img = pictureBox1.Image;
                    Bitmap bmp = new Bitmap(img.Width, img.Height);
                    Graphics gra = Graphics.FromImage(bmp);
                    gra.DrawImageUnscaled(img, new Point(0, 0));
                    gra.Dispose();

                    string deger = dataGridView1.SelectedCells[0].Value.ToString();
                    string ogrisim = "\\"+deger+".jpg";
                    currentcell = dataGridView1.CurrentCell.RowIndex;
                    if (currentcell!=0)
                    {
                        int iColumn = dataGridView1.CurrentCell.ColumnIndex;
                        int iRow = dataGridView1.CurrentCell.RowIndex;
                        if (iColumn == dataGridView1.Columns.Count - 1)
                            dataGridView1.CurrentCell = dataGridView1[0, iRow + 1];
                        else
                            dataGridView1.CurrentCell = dataGridView1[iColumn, iRow+1];

                    }
                    

                string belgelerim = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    bmp.Save(belgelerim + ogrisim, System.Drawing.Imaging.ImageFormat.Jpeg);
                    bmp.Dispose();
                    
                    int next = Convert.ToInt32(listView1.FocusedItem.Index) + 1;
                    
                    listView1.Items[next].Selected = true;

                    degisken = secilen.SubItems[sayac].Text;
                   
                    // focus değiştirmede kaldım focus değişecek yol değişecek aynı foto kaydedilmeyecek
                  

                }
            }
            catch { }
        }
        ListViewItem secilen; 
        string degisken;
        int sayac = 0;
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
             secilen = listView1.FocusedItem;//focusedıtem senin seçtiğin listview itemi alır
            
             degisken = secilen.SubItems[sayac].Text;

            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
           

            string alan = degisken;
            pictureBox1.ImageLocation = yol + "/" + alan;
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void pictureBox1_MouseDown_1(object sender, MouseEventArgs e)
        {
            // Seçimin başlangıç ​​noktası:
            if (e.Button == MouseButtons.Left)
            {
                
                _selecting = true;
                _selection = new Rectangle(new Point(e.X, e.Y), new Size());
            }

        }

        private void pictureBox1_MouseMove_1(object sender, MouseEventArgs e)
        {
            // Seçimin gerçek boyutunu güncelleyin:
            if (_selecting)
            {
                _selection.Width = e.X - _selection.X;
                _selection.Height = e.Y - _selection.Y;

                // Resim kutusunu yeniden çizin:
                pictureBox1.Refresh();
            }
        }

        private void pictureBox1_MouseUp_1(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && _selecting)
            {
               

                _selection.Width = e.X - _selection.X;
                _selection.Height = e.Y - _selection.Y;
                _selecting = false;
                

            }
          



        }

        private void pictureBox1_Paint_1(object sender, PaintEventArgs e)
        {
            if (_selecting)
            {
                //Geçerli seçimi gösteren bir dikdörtgen çizin
                Pen pen = Pens.GreenYellow;
                e.Graphics.DrawRectangle(pen, _selection);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Image img = pictureBox1.Image.Crop(_selection);
            pictureBox1.Image = img.Fit2PictureBox(pictureBox1);
            _selection.Width = 0;
            _selection.Height = 0;

            
            
              
            

        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            
        }
       
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
           
        }
    }
}

