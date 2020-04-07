using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Windows.Forms;
using ImageProcessor;
using Newtonsoft.Json;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace OCRAPI
{
    public partial class Form1 : Form
    {
        public string PdfPath { get; set; }

        private bool _isOrientationCorrect = false;
        public int ValueOrientation { get; set; } = -1;
        public int ValueLimitSizeFile { get; } = 1;

        private string _pathFileInput;
        private string _pathFileOutput;
        public string NameDirectoryOutput { get; } = "Output";
        public string ExtensionFile { get; set; }


        public int ValueResolutionInput { get; set; } = 200;
        public int ValueResolutionImage { get; set; } = -1;

        public Form1()
        {
            InitializeComponent();
            cmbLanguage.SelectedIndex = 18;//English
        }

        private string GetSelectedLanguage()
        {

            //https://ocr.space/OCRAPI#PostParameters

            //Czech = cze; Danish = dan; Dutch = dut; English = eng; Finnish = fin; French = fre; 
            //German = ger; Hungarian = hun; Italian = ita; Norwegian = nor; Polish = pol; Portuguese = por;
            //Spanish = spa; Swedish = swe; ChineseSimplified = chs; Greek = gre; Japanese = jpn; Russian = rus;
            //Turkish = tur; ChineseTraditional = cht; Korean = kor

            string strLang = "";
            switch (cmbLanguage.SelectedIndex)
            {
                case 0:
                    strLang = "ara";
                    break;

                case 1:
                    strLang = "chs";
                    break;

                case 2:
                    strLang = "cht";
                    break;
                case 3:
                    strLang = "cze";
                    break;
                case 4:
                    strLang = "dan";
                    break;
                case 5:
                    strLang = "dut";
                    break;
                case 6:
                    strLang = "eng";
                    break;
                case 7:
                    strLang = "fin";
                    break;
                case 8:
                    strLang = "fre";
                    break;
                case 9:
                    strLang = "ger";
                    break;
                case 10:
                    strLang = "gre";
                    break;
                case 11:
                    strLang = "hun";
                    break;
                case 12:
                    strLang = "jap";
                    break;
                case 13:
                    strLang = "kor";
                    break;
                case 14:
                    strLang = "nor";
                    break;
                case 15:
                    strLang = "pol";
                    break;
                case 16:
                    strLang = "por";
                    break;
                case 17:
                    strLang = "spa";
                    break;
                case 18:
                    strLang = "rus";
                    break;
                case 19:
                    strLang = "tur";
                    break;

            }
            return strLang;

        }

       private void button1_Click(object sender, EventArgs e)
        {
            PdfPath = _pathFileInput = ""; pictureBox.BackgroundImage = null;
            var fileDlg = new OpenFileDialog {Filter = @"jpeg, pdf files|*.jpg;*.JPG;*.pdf;*.PDF" };
            if (fileDlg.ShowDialog() == DialogResult.OK)
            {
                FileInfo fileInfo = new FileInfo(fileDlg.FileName);
                _pathFileInput = fileDlg.FileName;

                ExtensionFile = Path.GetExtension(_pathFileInput).ToLower();

                // если это PDF проводим конвертацию
                if (ExtensionFile == ".pdf")
                {
                    PdfPath = fileDlg.FileName;

                    // конвертация PDF в Jpeg
                    _pathFileInput = ConvertPdfToImage(PdfPath);
                }

                if (fileInfo.Length > ValueLimitSizeFile * 1024 * 1024)
                {
                    //Size limit depends: Free API 1 MB, PRO API 5 MB and more
                    MessageBox.Show(@"Image file size limit reached (1 MB Limit)");
                    return;
                }
                pictureBox.BackgroundImage = Image.FromFile(_pathFileInput);

                lblInfo.Text = @"Image loaded: "+ fileInfo.Name;
                lblInfo.BackColor = Color.LightGreen;
            }
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            PdfPath = _pathFileInput = "";
            pictureBox.BackgroundImage = null;
            OpenFileDialog fileDlg = new OpenFileDialog {Filter = @"pdf files|*.pdf;"};
            if (fileDlg.ShowDialog() == DialogResult.OK)
            {
                FileInfo fileInfo = new FileInfo(fileDlg.FileName);
                if (fileInfo.Length > ValueLimitSizeFile * 1024 * 1024 )
                {
                    //Size limit depends: Free API 1 MB, PRO API 5 MB and more
                    MessageBox.Show(@"PDF file size should not be larger than 1 Mb");
                    return;
                }
                PdfPath = fileDlg.FileName;

                // конвертация PDF в Jpeg
                _pathFileInput = ConvertPdfToImage(PdfPath);

                pictureBox.BackgroundImage = Image.FromFile(_pathFileInput);
                lblInfo.Text = @"PDF loaded: " + fileInfo.Name;
                lblInfo.BackColor = Color.LightSalmon;
            }
        }

        private byte[] ImageToBase64(Image image, System.Drawing.Imaging.ImageFormat format)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // Convert Image to byte[]
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();

                return imageBytes;
            }
        }

        private static byte[] CorrectImage(byte[] pImageInout, int pResolution, int pOrientation)
        {
            // Format is automatically detected though can be changed.
            Debug.Assert(pImageInout != null);
            using (MemoryStream inStream = new MemoryStream(pImageInout))
            {
                using (MemoryStream outStream = new MemoryStream())
                {
                    // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                    using (ImageFactory imageFactory = new ImageFactory(preserveExifData: true))
                    {
                        imageFactory.Load(inStream);
                        imageFactory.Quality(100);
                        //var arrPropertiesImages = imageFactory.ExifPropertyItems;
                        //imageFactory.AutoRotate();

                        var vResolutionCurrent = (int)imageFactory.Image.HorizontalResolution;

                        if (pResolution > 0 && vResolutionCurrent != pResolution)
                            imageFactory.Resolution(pResolution, pResolution);
                        
                        if (pOrientation > 0)
                            imageFactory.Rotate(-pOrientation);

                        imageFactory.Save(outStream);

                    }

                    return outStream.ToArray();
                }
            }
        }

        private void SaveImageOnDisk(byte[] pImageInput)
        {
            var pathDirectory = Path.GetDirectoryName(_pathFileInput);

            try
            {
                if (!Directory.Exists(pathDirectory)) throw new Exception("Директория не найдена: " + pathDirectory);
                Directory.CreateDirectory(pathDirectory + "\\" + NameDirectoryOutput);

                _pathFileOutput = pathDirectory + "\\" + NameDirectoryOutput + "\\" + Path.GetFileName(_pathFileInput);
                File.WriteAllBytes(_pathFileOutput, pImageInput);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new Exception("Error : SaveImageOnDisk: " + e.Message);
            }
        }

        string ConvertPdfToImage(string pathFilePdf)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(pathFilePdf);
            
            Image bmp = doc.SaveAsImage(0);
            var pathDirectoryImage = Path.GetDirectoryName(pathFilePdf);
            var nameFileImage = Path.GetFileNameWithoutExtension(pathFilePdf) + ".jpeg";
            bmp.Save(pathDirectoryImage + "\\" + nameFileImage, ImageFormat.Jpeg);
            return pathDirectoryImage + "\\" + nameFileImage;
        }

        /*void CheckImage()
        {
            // Instantiate the reader
            Debug.Assert(!string.IsNullOrEmpty(_pathFileInput));

            var nameExtension = Path.GetExtension(_pathFileInput);
            if (nameExtension != ".jpeg" && nameExtension != ".jpg") return;

            using (ExifReader reader = new ExifReader(_pathFileInput))
            {
                // Extract the tag data using the ExifTags enumeration
                //if (reader.GetTagValue<decimal>(ExifTags.Orientation,
                //    out var orientation))
                //{
                //    ValueOrientation = (int) orientation;
                //}

                DateTime datePictureTaken;
                if (reader.GetTagValue<DateTime>(ExifTags.DateTimeDigitized,
                    out datePictureTaken))
                {
                    // Do whatever is required with the extracted information
                    MessageBox.Show(this, string.Format("The picture was taken on {0}",
                        datePictureTaken), "Image information", MessageBoxButtons.OK);
                }
            }
        }*/

        private async void button2_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrEmpty(_pathFileInput) && string.IsNullOrEmpty(PdfPath))
                return;

            //pictureBoxNormilized.Text = "";

            button1.Enabled = false;
            button2.Enabled = false;
            btnPDF.Enabled = false;

            try
            {
                HttpClient httpClient = new HttpClient {Timeout = new TimeSpan(1, 1, 1)};
                

                MultipartFormDataContent form = new MultipartFormDataContent
                {
                    {new StringContent("a1aeb2404288957"), "apikey"},
                    {new StringContent(GetSelectedLanguage()), "language"},
                    {new StringContent("true"), "detectOrientation"}
                };


                byte[] imageData = null;
                if (string.IsNullOrEmpty(_pathFileInput) == false)
                {
                    imageData = File.ReadAllBytes(_pathFileInput);
                    form.Add(new ByteArrayContent(imageData, 0, imageData.Length), "image", "image.jpg");
                }
                else
                {
                    throw  new Exception("Отсутствует параметр входящей строки");
                }

                HttpResponseMessage response = await httpClient.PostAsync("https://api.ocr.space/Parse/Image", form);

                string strContent = await response.Content.ReadAsStringAsync();

                RootObject ocrResult = JsonConvert.DeserializeObject<RootObject>(strContent);

                // UPDATE добавить проверку распознанного и предупредить, если не распозналось (сохранить в отдельную папку)

                if (ocrResult.OcrExitCode == 1)
                {
                    for (var i = 0; i < ocrResult.ParsedResults.Count() ; i++)
                    {
                        var orientation = ocrResult.ParsedResults[i].TextOrientation;
                        int.TryParse(orientation, out var vRotateAngle);

                        // редактируем изображение на основе полученных данных
                        // поворот и смена разрешения (dpi)
                        imageData = CorrectImage(imageData, ValueResolutionInput, vRotateAngle);
                        
                        // сохраняем полученную картинку
                        SaveImageOnDisk(imageData);

                        pictureBoxNormilized.Text += ocrResult.ParsedResults[i].ParsedText ;
                        pictureBoxNormilized.BackgroundImage = Image.FromFile(_pathFileOutput);
                    }
                }
                else
                {
                    MessageBox.Show(@"OCR Error: " + strContent);
                }

            }
            catch (Exception exception)
            {
                MessageBox.Show(@"Error: " + exception.Message);
            }

            button1.Enabled = true;
            button2.Enabled = true;
            btnPDF.Enabled = true;
        }

        private void numericDpi_ValueChanged(object sender, EventArgs e) => ValueResolutionInput = (int) numericDpi.Value;
    }
}



