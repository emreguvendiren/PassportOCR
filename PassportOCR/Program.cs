using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;
using Size = OpenCvSharp.Size;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;

using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Color = System.Drawing.Color;
using Rectangle = System.Drawing.Rectangle;
using System.Drawing.Imaging;
using System.Diagnostics;

namespace PassportScanner
{
    class Program
    {
        
        static bool kontrol = false;
        static void Main(string[] args)
        {

            List<String> pass = new List<string>();
            try
            {

                string path5 = "Enter the pdf's path";
                string ocr = GetPassportInformacition(path5);
                pass.Add(ocr);
                if (pass.Count < 1 || pass[0] == "")
                {
                    kontrol = true;
                    string ocr2 = GetPassportInformacition(path5);
                    pass.Add(ocr2);

                }
                if (pass.Count < 1 || pass[0].Equals(""))
                {
                    Console.WriteLine("OKUNAMADI");
                }
                foreach (var item in pass)
                {
                    if (item[0] == 'P' && item[1] == '<')
                    {
                        Console.WriteLine(item);
                    }

                }
            }

            catch (Exception e)
            {

                Console.WriteLine("Beklenmeyen bir hata oluştu. Hata Mesajı:  " + e.Message);
            }

            Console.Read();


        }
        public static void TrimImage(string path)
        {
            int threshhold = 200;


            int topOffset = 0;
            int bottomOffset = 0;
            int leftOffset = 0;
            int rightOffset = 0;
            Bitmap img = new Bitmap(path);


            bool foundColor = false;
            // Get left bounds to crop
            for (int x = 1; x < img.Width && foundColor == false; x++)
            {
                for (int y = 1; y < img.Height && foundColor == false; y++)
                {
                    Color color = img.GetPixel(x, y);
                    if (color.R < threshhold || color.G < threshhold || color.B < threshhold)
                        foundColor = true;
                }
                leftOffset += 1;
            }


            foundColor = false;
            // Get top bounds to crop
            for (int y = 1; y < img.Height && foundColor == false; y++)
            {
                for (int x = 1; x < img.Width && foundColor == false; x++)
                {
                    Color color = img.GetPixel(x, y);
                    if (color.R < threshhold || color.G < threshhold || color.B < threshhold)
                        foundColor = true;
                }
                topOffset += 1;
            }


            foundColor = false;
            // Get right bounds to crop
            for (int x = img.Width - 1; x >= 1 && foundColor == false; x--)
            {
                for (int y = 1; y < img.Height && foundColor == false; y++)
                {
                    Color color = img.GetPixel(x, y);
                    if (color.R < threshhold || color.G < threshhold || color.B < threshhold)
                        foundColor = true;
                }
                rightOffset += 1;
            }


            foundColor = false;
            // Get bottom bounds to crop
            for (int y = img.Height - 1; y >= 1 && foundColor == false; y--)
            {
                for (int x = 1; x < img.Width && foundColor == false; x++)
                {
                    Color color = img.GetPixel(x, y);
                    if (color.R < threshhold || color.G < threshhold || color.B < threshhold)
                        foundColor = true;
                }
                bottomOffset += 1;
            }



            Bitmap croppedBitmap = new Bitmap(img);
            croppedBitmap = croppedBitmap.Clone(
                            new Rectangle(leftOffset, topOffset, img.Width - leftOffset - rightOffset, img.Height - topOffset - bottomOffset),
                            System.Drawing.Imaging.PixelFormat.DontCare);
            /*
             croppedBitmap = croppedBitmap.Clone(
                            new Rectangle(leftOffset , topOffset , img.Width - leftOffset - rightOffset + 6, img.Height - topOffset - bottomOffset + 6),
                            System.Drawing.Imaging.PixelFormat.DontCare);
             */

            string realPath = Regex.Replace(path, ".png", String.Empty);
            croppedBitmap.Save(realPath + "_white" + ".png", ImageFormat.Png);
        }
        public static string GetPassportInformacition(string Path)
        {
            List<String> Paths = new List<string>();

            #region pathin sonundaki .pdf i silme
            string notPNG = Path;
            string realPath = Regex.Replace(notPNG, ".pdf", String.Empty);



            if (kontrol == true)
            {
                realPath = realPath + "2";
            }
            #endregion
            #region PDF TO IMAGE
            var doc = new Document(Path);
            int sayc = 0;
            foreach (var item in doc.Pages)
            {

                sayc++;
            }
            if (sayc > 1)
            {

                using (FileStream img = new FileStream(realPath + ".png", FileMode.Create))
                {
                    Resolution resolution = new Resolution(280);
                    PngDevice pngDevice = new PngDevice(resolution);
                    pngDevice.Process(doc.Pages[2], img);
                    img.Close();
                }
            }
            else
            {
                using (FileStream img = new FileStream(realPath + ".png", FileMode.Create))
                {
                    Resolution resolution = new Resolution(280);
                    PngDevice pngDevice = new PngDevice(resolution);
                    pngDevice.Process(doc.Pages[1], img);
                    img.Close();
                }

            }



            TrimImage(realPath + ".png");



            #endregion
            #region Projenin bulundugu dizinin pathini alma
            string datapath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName;
            string asd = Directory.GetCurrentDirectory();

            string deletePath = datapath;

            #endregion


            Guid obj = Guid.NewGuid();

            string pathimg = realPath + "_white";
            string writeimg = datapath;

            Rect bigRect;
            Mat rgb;

            int sayac = 0;
            string path = realPath + "_white";
            List<String> PassportNumbers = new List<string>();
            List<String> names = new List<string>();
            for (int a = 0; a < 4; a++)
            {
                try
                {
                    #region Fotografi kesme ve iyilestirme
                    var bitmapCrop = new Mat(path + ".png");
                    if (kontrol == false)
                    {
                        if (bitmapCrop.Width > 1800)
                        {
                            Bitmap redCut = new Bitmap(path + ".png");
                            Bitmap redCrop = redCut.Clone(new System.Drawing.Rectangle(0, 1300, bitmapCrop.Width - 0, bitmapCrop.Height - 1300), redCut.PixelFormat);

                            redCrop.Save(pathimg + "out" + ".png");

                        }
                        else
                        {
                            //int yukseklik = bitmapCrop.Height * 3 / 4;
                            //int borderHeight = bitmapCrop.Height / 4;
                            Bitmap redCut = new Bitmap(path + ".png");
                            Bitmap redCrop = redCut.Clone(new System.Drawing.Rectangle(0, 1300, bitmapCrop.Width, bitmapCrop.Height - 1300), redCut.PixelFormat);

                            redCrop.Save(pathimg + "out" + ".png");

                        }
                    }
                    else
                    {
                        if (bitmapCrop.Width > 1800)
                        {
                            Bitmap redCut = new Bitmap(path + ".png");
                            Bitmap redCrop = redCut.Clone(new System.Drawing.Rectangle(100, 1300, bitmapCrop.Width - 100, bitmapCrop.Height - 1300), redCut.PixelFormat);

                            redCrop.Save(pathimg + "out" + ".png");

                        }
                        else
                        {
                            //int yukseklik = bitmapCrop.Height * 3 / 4;
                            //int borderHeight = bitmapCrop.Height / 4;
                            Bitmap redCut = new Bitmap(path + ".png");
                            Bitmap redCrop = redCut.Clone(new System.Drawing.Rectangle(75, 1300, bitmapCrop.Width - 75, bitmapCrop.Height - 1300), redCut.PixelFormat);

                            redCrop.Save(pathimg + "out" + ".png");

                        }
                    }

                    #endregion


                    Mat src = new Mat(pathimg + "out" + ".Png", ImreadModes.Color);



                    rgb = new Mat(pathimg + "out" + ".png", ImreadModes.Color);

                    Mat gray = new Mat();
                    Cv2.CvtColor(rgb, gray, ColorConversionCodes.BGR2GRAY);


                    Mat grad = new Mat();

                    Mat morphKernel = new Mat();
                    if (a < 4)
                    {
                        morphKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(10, 10));
                    }
                    else
                    {
                        morphKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(10, 10));
                    }

                    Cv2.MorphologyEx(gray, grad, MorphTypes.Gradient, morphKernel);

                    Mat bw = new Mat();

                    Cv2.Threshold(grad, bw, 0.0, 255.0, ThresholdTypes.Binary | ThresholdTypes.Otsu | ThresholdTypes.Tozero);


                    Mat connected = new Mat();

                    morphKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(25, 5));

                    Cv2.MorphologyEx(bw, connected, MorphTypes.Close, morphKernel);

                    Mat mask = Mat.Zeros(bw.Size(), MatType.CV_8UC1);




                    OpenCvSharp.Point[][] contours;

                    HierarchyIndex[] hierarchy;


                    Cv2.FindContours(connected, out contours, out hierarchy, RetrievalModes.CComp, ContourApproximationModes.ApproxSimple, new OpenCvSharp.Point(0, 0));

                    Mat newImage = new Mat();
                    rgb.CopyTo(newImage);
                    


                    List<Rect> mrz = new List<Rect>();
                    List<Rect> mrz2 = new List<Rect>();
                    
                    string TessData = datapath + @"\tessdata";

                    double r = 0;

                    for (int i = 0; i < hierarchy.Length; i++)
                    {

                        Rect rect = Cv2.BoundingRect(contours[i]);

                        try
                        {
                            if (rect.Height > 30 && rect.Width > 1000)
                            {
                                try
                                {
                                    mrz2.Add(rect);
                                    Cv2.Rectangle(rgb, rect, new Scalar(0, 255, 0), 1);
                                }
                                catch (Exception)
                                {}
                            }
                        }
                        catch (Exception)
                        {


                        }
                        r = rect.Height > 0 ? (double)(rect.Width / rect.Height) : 0;

                        if ((rect.Width > connected.Cols * 0.4) &&
                            (r > 15) &&
                            (r < 39))
                        {
                            mrz.Add(rect);

                            Cv2.Rectangle(rgb, rect, new Scalar(0, 255, 0), 1);

                        }
                        else
                        {

                            Cv2.Rectangle(rgb, rect, new Scalar(0, 0, 255), 1);

                        }


                    }

                    Mat testmat = new Mat();

                    bigRect = new Rect();


                    if (mrz.Count == 2)
                    {
                        Rect max = Rect.Union(mrz[0], mrz[1]);
                        bigRect = max;
                        Cv2.Rectangle(rgb, max, new Scalar(500, 200, 0), 1);
                        int saveControl = 0;
                        for (int i = 0; i < 2; i++)
                        {
                            var tesseract2 = OpenCvSharp.Text.OCRTesseract.Create(TessData, "mrz", null, 3, 6);

                            tesseract2.SetWhiteList("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ<");
                            tesseract2.Run(rgb[mrz[i]], out var outputText2, out var componentRects2, out var componentTexts2, out var componentConfidences2);
                            string ocredtext2 = outputText2.Trim().Replace(" ", string.Empty);
                            string newStr2 = ocredtext2.Trim().Replace("\n", string.Empty);
                            saveControl = newStr2.IndexOf('<');
                            names.Add(newStr2);
                        }
                        
                        if (saveControl !=-1)
                        {
                            Bitmap bmp2 = new Bitmap(pathimg + "out" + ".Png");
                            bmp2.Save(writeimg);
                        }

                    }
                    else if (mrz.Count == 3)
                    {
                        Rect max2 = Rect.Union(mrz[0], mrz[1]);
                        Rect max3 = Rect.Union(mrz[2], max2);
                        bigRect = max3;
                        Cv2.Rectangle(rgb, max3, new Scalar(500, 200, 0), 1);
                        int saveControl = 0;
                        for (int i = 0; i < 3; i++)
                        {
                            var tesseract2 = OpenCvSharp.Text.OCRTesseract.Create(TessData, "mrz", null, 3, 6);

                            tesseract2.SetWhiteList("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ<");
                            tesseract2.Run(rgb[mrz[i]], out var outputText2, out var componentRects2, out var componentTexts2, out var componentConfidences2);
                            string ocredtext2 = outputText2.Trim().Replace(" ", string.Empty);
                            string newStr2 = ocredtext2.Trim().Replace("\n", string.Empty);
                            saveControl = newStr2.IndexOf('<');
                            names.Add(newStr2);
                        }

                        if (saveControl != -1)
                        {
                            Bitmap bmp2 = new Bitmap(pathimg + "out" + ".Png");
                            bmp2.Save(writeimg);
                        }
                    }
                    else if (mrz2.Count > 1)
                    {
                        int saveControl = 0;
                        for (int i = 0; i < mrz2.Count; i++)
                        {
                            try
                            {

                                var tesseract2 = OpenCvSharp.Text.OCRTesseract.Create(TessData, "mrz", null, 3, 6);

                                tesseract2.SetWhiteList("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ<");
                                tesseract2.Run(rgb[mrz[i]], out var outputText2, out var componentRects2, out var componentTexts2, out var componentConfidences2);
                                string ocredtext2 = outputText2.Trim().Replace(" ", string.Empty);
                                string newStr2 = ocredtext2.Trim().Replace("\n", string.Empty);
                                saveControl = newStr2.IndexOf('<');
                                names.Add(newStr2);
                            }
                            catch (Exception)
                            {
                                continue;

                            }
                        }
                        if (saveControl != -1)
                        {
                            Bitmap bmp2 = new Bitmap(pathimg + "out" + ".Png");
                            bmp2.Save(writeimg);
                        }

                    }




                    pathimg = path;
                    #region Fotografi 90 derece dondurme
                    Bitmap bmp = new Bitmap(path + ".png");
                    path += sayac.ToString();
                    bmp.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipNone);
                    bmp.Save(path + ".png");
                    #endregion
                    Paths.Add(pathimg);
                    pathimg = pathimg + sayac.ToString();

                    sayac++;


                }
                catch (Exception )
                {


                    pathimg = path;
                    #region Fotografi 90 derece dondurme
                    Bitmap bmp = new Bitmap(path + ".png");
                    path += sayac.ToString();
                    bmp.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipNone);
                    bmp.Save(path + ".png");
                    #endregion
                    Paths.Add(pathimg);
                    pathimg = pathimg + sayac.ToString();


                    sayac++;
                }

            }

            List<String> passportNum = new List<String>();

            string passportInformation= "";
            string ilk = "";
            string son = "";
            foreach (var item in names)
            {
                int say = 0;
                for (int i = 0; i < item.Length; i++)
                {
                    if (item[i] == '<') { say++; }
                }
                if (say > 6)
                {
                    passportInformation = item;
                    ilk = item;
                    passportNum.Add(item);
                }
            }
            string[] bol = passportInformation.Split('<');
            string name = bol[1];
            string region = "";
            for (int i = 0; i < 3; i++)
            {
                region += name[i];
            }
            foreach (var item in names)
            {
                int control = item.IndexOf(region);
                if (control != -1 && ilk != item)
                {
                    son = item;
                }
            }
            string info = ilk + son;
            return info;

        }



    }
}