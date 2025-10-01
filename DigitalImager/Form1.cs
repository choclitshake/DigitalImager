using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WebCamLib;

namespace DigitalImager
{
    public partial class Form1 : Form
    {
        private Panel previewPanel;
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        private enum EffectMode { None, Copy, Greyscale, Invert, Sepia, Histogram, Subtract, Smooth, Gaussian, Sharpen, MeanRemoval, EmbossLaplacian, EmbossHorzVert, EmbossAllDir, EmbossLossy, EmbossHorizontal, EmbossVertical }
        private EffectMode currentEffect = EffectMode.None;
        private bool isWebcamRunning = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void loadImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                turnOffToolStripMenuItem_Click(sender, e);
            }
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select an Image";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using (var original = new Bitmap(ofd.FileName))
                    {
                        pictureBox1.Image?.Dispose();
                        pictureBox1.Image = ResizeAndCrop(original, pictureBox1.Width, pictureBox1.Height);
                    }
                }
            }
        }

        private void loadBackgroundToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select a Background Image";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using (var original = new Bitmap(ofd.FileName))
                    {
                        pictureBox2.Image?.Dispose();
                        pictureBox2.Image = ResizeAndCrop(original, pictureBox1.Width, pictureBox1.Height);
                    }
                }
            }
        }

        private void saveImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox3.Image != null)
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Title = "Save Image As";
                    sfd.Filter = "PNG Image|*.png|JPEG Image|*.jpg|Bitmap Image|*.bmp";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        ImageFormat format = ImageFormat.Png;

                        switch (System.IO.Path.GetExtension(sfd.FileName).ToLower())
                        {
                            case ".jpg":
                            case ".jpeg":
                                format = ImageFormat.Jpeg;
                                break;
                            case ".bmp":
                                format = ImageFormat.Bmp;
                                break;
                        }

                        pictureBox3.Image.Save(sfd.FileName, format);
                        MessageBox.Show("Image saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("No result image yet!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void basicCopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (videoSource != null && videoSource.IsRunning)
                currentEffect = EffectMode.Copy;
            else
                ApplyEffectToImage(EffectMode.Copy);
        }

        private void greyscaleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Greyscale;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }

                Bitmap source = new Bitmap(pictureBox1.Image);
                Bitmap result = ApplyGreyscale(source);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                source.Dispose();
            }
        }

        private void colorInversionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Invert;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }

                Bitmap source = new Bitmap(pictureBox1.Image);
                Bitmap result = ApplyInvert(source);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                source.Dispose();
            }
        }

        private void sepiaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Sepia;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }

                Bitmap source = new Bitmap(pictureBox1.Image);
                Bitmap result = ApplySepia(source);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                source.Dispose();
            }
        }

        private void histogramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Histogram;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }

                Bitmap source = new Bitmap(pictureBox1.Image);
                Bitmap result = GenerateHistogram(source);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                source.Dispose();
            }
        }

        private void subtractToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                if (pictureBox2.Image == null)
                {
                    MessageBox.Show("Load a background image first.");
                    return;
                }
                currentEffect = EffectMode.Subtract;
            }
            else
            {
                if (pictureBox1.Image == null || pictureBox2.Image == null)
                {
                    MessageBox.Show("Load both main and background images first.");
                    return;
                }

                Bitmap source = new Bitmap(pictureBox1.Image);
                Bitmap background = new Bitmap(pictureBox2.Image);
                Bitmap result = ApplySubtract(source, background);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                source.Dispose();
                background.Dispose();
            }
        }

        private void smoothToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Smooth;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.Smooth(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void gaussianBlurToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Gaussian;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.GaussianBlur(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void sharpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.Sharpen;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.Sharpen(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void meanRemovalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.MeanRemoval;
            }
            else
            {
                if (pictureBox1.Image == null)
                {
                    MessageBox.Show("Load an image first.");
                    return;
                }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.MeanRemoval(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossLaplacianToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossLaplacian;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossLaplacian(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossHorzVertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossHorzVert;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossHorzVert(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossAllDirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossAllDir;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossAllDir(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossLossyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossLossy;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossLossy(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossHorizontal;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossHorizontal(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void embossVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isWebcamRunning)
            {
                currentEffect = EffectMode.EmbossVertical;
            }
            else
            {
                if (pictureBox1.Image == null) { MessageBox.Show("Load an image first."); return; }
                Bitmap src = new Bitmap(pictureBox1.Image);
                Bitmap result = new Bitmap(src);
                BitmapFilter.EmbossVertical(result);
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = result;
                src.Dispose();
            }
        }

        private void turnOnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isWebcamRunning = true;

            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                MessageBox.Show("No webcam found!");
                return;
            }

            videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);

            if (videoSource.VideoCapabilities.Length > 0)
            {
                var lowRes = videoSource.VideoCapabilities
                                        .OrderBy(v => v.FrameSize.Width * v.FrameSize.Height)
                                        .First();
                videoSource.VideoResolution = lowRes;
            }
            
            try
            {
                videoSource.DesiredFrameRate = 15; 
            }
            catch { }

            videoSource.NewFrame += VideoSource_NewFrame;
            videoSource.Start();
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap rawFrame = null;
            try
            {
                rawFrame = (Bitmap)eventArgs.Frame.Clone();
            }
            catch
            {
                return;
            }

            if (pictureBox1.IsHandleCreated && !pictureBox1.IsDisposed)
            {
                var frameForPb1 = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox1.Width, pictureBox1.Height);
                pictureBox1.BeginInvoke(new MethodInvoker(() =>
                {
                    pictureBox1.Image?.Dispose();
                    pictureBox1.Image = frameForPb1;
                }));
            }

            Bitmap processedFrame = null;

            switch (currentEffect)
            {
                case EffectMode.Copy:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    break;
                case EffectMode.Greyscale: 
                    processedFrame = ResizeAndCrop(ApplyGreyscale((Bitmap)rawFrame.Clone()), pictureBox3.Width, pictureBox3.Height);
                    break;
                case EffectMode.Invert:
                    processedFrame = ResizeAndCrop(ApplyInvert((Bitmap)rawFrame.Clone()), pictureBox3.Width, pictureBox3.Height);
                    break;
                case EffectMode.Sepia:
                    processedFrame = ResizeAndCrop(ApplySepia((Bitmap)rawFrame.Clone()), pictureBox3.Width, pictureBox3.Height);
                    break;
                case EffectMode.Subtract:
                    if (pictureBox2.Image != null)
                    {
                        using (Bitmap bg = new Bitmap(pictureBox2.Image, rawFrame.Size))
                        {
                            processedFrame = ResizeAndCrop(ApplySubtract((Bitmap)rawFrame.Clone(), new Bitmap(pictureBox2.Image)), pictureBox3.Width, pictureBox3.Height);
                        }
                    }
                    break;
                case EffectMode.Histogram:
                    processedFrame = ResizeAndCrop(GenerateHistogram((Bitmap)rawFrame.Clone()), pictureBox3.Width, pictureBox3.Height);
                    break;
                case EffectMode.Smooth:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.Smooth(processedFrame);
                    break;
                case EffectMode.Gaussian:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.GaussianBlur(processedFrame);
                    break;
                case EffectMode.Sharpen:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.Sharpen(processedFrame);
                    break;
                case EffectMode.MeanRemoval:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.MeanRemoval(processedFrame);
                    break;
                case EffectMode.EmbossLaplacian:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossLaplacian(processedFrame);
                    break;
                case EffectMode.EmbossHorzVert:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossHorzVert(processedFrame);
                    break;
                case EffectMode.EmbossAllDir:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossAllDir(processedFrame);
                    break;
                case EffectMode.EmbossLossy:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossLossy(processedFrame);
                    break;
                case EffectMode.EmbossHorizontal:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossHorizontal(processedFrame);
                    break;
                case EffectMode.EmbossVertical:
                    processedFrame = ResizeAndCrop((Bitmap)rawFrame.Clone(), pictureBox3.Width, pictureBox3.Height);
                    BitmapFilter.EmbossVertical(processedFrame);
                    break;
            }

            if (pictureBox3.IsHandleCreated && !pictureBox3.IsDisposed)
            {
                pictureBox3.BeginInvoke(new MethodInvoker(() =>
                {
                    pictureBox3.Image?.Dispose();
                    pictureBox3.Image = processedFrame;
                }));
            }
            else
            {
                processedFrame?.Dispose();
            }

            rawFrame.Dispose();
        }

        private void turnOffToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isWebcamRunning = false;

            if (videoSource != null)
            {
                if (videoSource.IsRunning)
                {
                    videoSource.SignalToStop();
                    videoSource.WaitForStop(); 
                }
                videoSource.NewFrame -= VideoSource_NewFrame;
                videoSource = null;
            }

            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
            }

            if (pictureBox3.Image != null)
            {
                pictureBox3.Image.Dispose();
                pictureBox3.Image = null;
            }

            currentEffect = EffectMode.None;
        }

        private void ApplyEffectToImage(EffectMode effect)
        {
            if (pictureBox1.Image == null)
            {
                MessageBox.Show("Load an image first.");
                return;
            }

            Bitmap source = new Bitmap(pictureBox1.Image);
            switch (effect)
            {
                case EffectMode.Greyscale:
                    ApplyGreyscale(source);
                    break;
                case EffectMode.Invert:
                    ApplyInvert(source);
                    break;
                case EffectMode.Sepia:
                    ApplySepia(source);
                    break;
            }

            pictureBox3.Image?.Dispose();
            pictureBox3.Image = source;
        }

        private Bitmap ApplyGreyscale(Bitmap source)
        {
            Bitmap bmp = new Bitmap(source);
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            BitmapData data =
                bmp.LockBits(rect, ImageLockMode.ReadWrite, bmp.PixelFormat);

            int bytesPerPixel = Image.GetPixelFormatSize(bmp.PixelFormat) / 8;
            int byteCount = data.Stride * bmp.Height;
            byte[] pixels = new byte[byteCount];
            Marshal.Copy(data.Scan0, pixels, 0, byteCount);

            for (int y = 0; y < bmp.Height; y++)
            {
                int offset = y * data.Stride;
                for (int x = 0; x < bmp.Width; x++)
                {
                    int index = offset + x * bytesPerPixel;

                    byte r = pixels[index + 2];
                    byte g = pixels[index + 1];
                    byte b = pixels[index];

                    byte gray = (byte)((r + g + b) / 3);

                    pixels[index] = gray;     // b
                    pixels[index + 1] = gray; // g
                    pixels[index + 2] = gray; // r
                }
            }

            Marshal.Copy(pixels, 0, data.Scan0, byteCount);
            bmp.UnlockBits(data);

            return bmp;
        }

        private Bitmap ApplyInvert(Bitmap source)
        {
            Bitmap bmp = new Bitmap(source);
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            var data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadWrite, bmp.PixelFormat);

            int bpp = Image.GetPixelFormatSize(bmp.PixelFormat) / 8;
            int byteCount = data.Stride * bmp.Height;
            byte[] pixels = new byte[byteCount];
            Marshal.Copy(data.Scan0, pixels, 0, byteCount);

            for (int i = 0; i < pixels.Length; i += bpp)
            {
                pixels[i] = (byte)(255 - pixels[i]);     // b
                pixels[i + 1] = (byte)(255 - pixels[i + 1]); // g
                pixels[i + 2] = (byte)(255 - pixels[i + 2]); // r
            }

            Marshal.Copy(pixels, 0, data.Scan0, byteCount);
            bmp.UnlockBits(data);

            return bmp;
        }

        private Bitmap ApplySepia(Bitmap source)
        {
            Bitmap bmp = new Bitmap(source);
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            var data = bmp.LockBits(rect, ImageLockMode.ReadWrite, bmp.PixelFormat);

            int bpp = Image.GetPixelFormatSize(bmp.PixelFormat) / 8;
            int byteCount = data.Stride * bmp.Height;
            byte[] pixels = new byte[byteCount];
            Marshal.Copy(data.Scan0, pixels, 0, byteCount);

            for (int i = 0; i < pixels.Length; i += bpp)
            {
                byte b = pixels[i];
                byte g = pixels[i + 1];
                byte r = pixels[i + 2];

                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                pixels[i + 2] = (byte)Math.Min(255, tr); // r
                pixels[i + 1] = (byte)Math.Min(255, tg); // g
                pixels[i] = (byte)Math.Min(255, tb); // b
            }

            Marshal.Copy(pixels, 0, data.Scan0, byteCount);
            bmp.UnlockBits(data);

            return bmp;
        }

        private Bitmap GenerateHistogram(Bitmap frame)
        {
            Bitmap source = (Bitmap)frame.Clone();
            int[] histogram = new int[256];

            Rectangle rect = new Rectangle(0, 0, source.Width, source.Height);
            BitmapData data = source.LockBits(rect,
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);

            unsafe
            {
                byte* ptr = (byte*)data.Scan0;
                int stride = data.Stride;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* row = ptr + (y * stride);
                    for (int x = 0; x < source.Width; x++)
                    {
                        byte b = row[x * 3];
                        byte g = row[x * 3 + 1];
                        byte r = row[x * 3 + 2];

                        int gray = (r + g + b) / 3;
                        histogram[gray]++;
                    }
                }
            }

            source.UnlockBits(data);
            source.Dispose();

            int max = histogram.Max();
            Bitmap histImage = new Bitmap(256, 200);

            using (Graphics g = Graphics.FromImage(histImage))
            {
                g.Clear(Color.White);
                for (int i = 0; i < 256; i++)
                {
                    int height = (int)(histogram[i] * 200f / max);
                    g.DrawLine(Pens.Black, new Point(i, 200), new Point(i, 200 - height));
                }
            }

            return histImage;
        }

        private Bitmap ApplySubtract(Bitmap sourceImage, Bitmap backgroundImage)
        {
            if (sourceImage == null)
                throw new ArgumentNullException(nameof(sourceImage));

            int width = sourceImage.Width;
            int height = sourceImage.Height;

            if (backgroundImage == null)
                return new Bitmap(sourceImage);

            Bitmap sourceCopy = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(sourceCopy))
                g.DrawImage(sourceImage, 0, 0, width, height);

            Bitmap backgroundCopy = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(backgroundCopy))
                g.DrawImage(backgroundImage, 0, 0, width, height);

            Rectangle rect = new Rectangle(0, 0, width, height);

            BitmapData sourceData = sourceCopy.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            BitmapData backgroundData = backgroundCopy.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            int stride = sourceData.Stride;
            int byteCount = stride * height;

            byte[] sourceBytes = new byte[byteCount];
            byte[] backgroundBytes = new byte[byteCount];
            Marshal.Copy(sourceData.Scan0, sourceBytes, 0, byteCount);
            Marshal.Copy(backgroundData.Scan0, backgroundBytes, 0, byteCount);

            sourceCopy.UnlockBits(sourceData);
            backgroundCopy.UnlockBits(backgroundData);

            byte[] resultBytes = new byte[byteCount];

            int threshold = 60;

            for (int y = 0; y < height; y++)
            {
                int rowOffset = y * stride;
                for (int x = 0; x < width; x++)
                {
                    int index = rowOffset + x * 3;

                    byte blue = sourceBytes[index];
                    byte green = sourceBytes[index + 1];
                    byte red = sourceBytes[index + 2];

                    if (green > red * 1.3 && green > blue * 1.3 && green > 50)
                    {
                        resultBytes[index] = backgroundBytes[index];
                        resultBytes[index + 1] = backgroundBytes[index + 1];
                        resultBytes[index + 2] = backgroundBytes[index + 2];
                    }
                    else
                    {
                        resultBytes[index] = blue;
                        resultBytes[index + 1] = green;
                        resultBytes[index + 2] = red;
                    }
                }
            }

            Bitmap resultImage = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            BitmapData resultData = resultImage.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            Marshal.Copy(resultBytes, 0, resultData.Scan0, byteCount);
            resultImage.UnlockBits(resultData);

            sourceCopy.Dispose();
            backgroundCopy.Dispose();

            return resultImage;
        }
        private Bitmap ResizeAndCrop(Bitmap input, int targetWidth, int targetHeight)
        {
            float ratioW = (float)targetWidth / input.Width;
            float ratioH = (float)targetHeight / input.Height;
            float scale = Math.Max(ratioW, ratioH); 

            int scaledWidth = (int)(input.Width * scale);
            int scaledHeight = (int)(input.Height * scale);

            Bitmap scaled = new Bitmap(scaledWidth, scaledHeight);
            using (Graphics g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(input, new Rectangle(0, 0, scaledWidth, scaledHeight));
            }

            int offsetX = (scaledWidth - targetWidth) / 2;
            int offsetY = (scaledHeight - targetHeight) / 2;

            Rectangle cropRect = new Rectangle(offsetX, offsetY, targetWidth, targetHeight);
            Bitmap cropped = new Bitmap(targetWidth, targetHeight);
            using (Graphics g = Graphics.FromImage(cropped))
            {
                g.DrawImage(scaled, new Rectangle(0, 0, targetWidth, targetHeight), cropRect, GraphicsUnit.Pixel);
            }

            scaled.Dispose();
            return cropped;
        }
    }

    public class ConvMatrix
    {
        public int TopLeft = 0, TopMid = 0, TopRight = 0;
        public int MidLeft = 0, Pixel = 0, MidRight = 0;
        public int BottomLeft = 0, BottomMid = 0, BottomRight = 0;
        public int Factor = 1;
        public int Offset = 0;

        public void SetAll(int nVal)
        {
            TopLeft = TopMid = TopRight = MidLeft = Pixel = MidRight =
                      BottomLeft = BottomMid = BottomRight = nVal;
        }
    }

    public static class BitmapFilter
    {
        public static bool Conv3x3(Bitmap b, ConvMatrix m)
        {
            if (m.Factor == 0) return false;

            Bitmap bSrc = (Bitmap)b.Clone();

            BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height),
                                ImageLockMode.ReadWrite,
                                PixelFormat.Format24bppRgb);

            BitmapData bmSrc = bSrc.LockBits(new Rectangle(0, 0, bSrc.Width, bSrc.Height),
                                 ImageLockMode.ReadWrite,
                                 PixelFormat.Format24bppRgb);

            int stride = bmData.Stride;
            int stride2 = stride * 2;

            IntPtr Scan0 = bmData.Scan0;
            IntPtr SrcScan0 = bmSrc.Scan0;

            unsafe
            {
                byte* p = (byte*)(void*)Scan0;
                byte* pSrc = (byte*)(void*)SrcScan0;

                int nOffset = stride - b.Width * 3;
                int nWidth = b.Width - 2;
                int nHeight = b.Height - 2;

                int nPixel;

                for (int y = 0; y < nHeight; ++y)
                {
                    for (int x = 0; x < nWidth; ++x)
                    {
                        nPixel = (((pSrc[2] * m.TopLeft) +
                                   (pSrc[5] * m.TopMid) +
                                   (pSrc[8] * m.TopRight) +
                                   (pSrc[2 + stride] * m.MidLeft) +
                                   (pSrc[5 + stride] * m.Pixel) +
                                   (pSrc[8 + stride] * m.MidRight) +
                                   (pSrc[2 + stride2] * m.BottomLeft) +
                                   (pSrc[5 + stride2] * m.BottomMid) +
                                   (pSrc[8 + stride2] * m.BottomRight))
                                   / m.Factor) + m.Offset;

                        if (nPixel < 0) nPixel = 0;
                        if (nPixel > 255) nPixel = 255;
                        p[5 + stride] = (byte)nPixel;

                        nPixel = (((pSrc[1] * m.TopLeft) +
                                   (pSrc[4] * m.TopMid) +
                                   (pSrc[7] * m.TopRight) +
                                   (pSrc[1 + stride] * m.MidLeft) +
                                   (pSrc[4 + stride] * m.Pixel) +
                                   (pSrc[7 + stride] * m.MidRight) +
                                   (pSrc[1 + stride2] * m.BottomLeft) +
                                   (pSrc[4 + stride2] * m.BottomMid) +
                                   (pSrc[7 + stride2] * m.BottomRight))
                                   / m.Factor) + m.Offset;

                        if (nPixel < 0) nPixel = 0;
                        if (nPixel > 255) nPixel = 255;
                        p[4 + stride] = (byte)nPixel;

                        nPixel = (((pSrc[0] * m.TopLeft) +
                                   (pSrc[3] * m.TopMid) +
                                   (pSrc[6] * m.TopRight) +
                                   (pSrc[0 + stride] * m.MidLeft) +
                                   (pSrc[3 + stride] * m.Pixel) +
                                   (pSrc[6 + stride] * m.MidRight) +
                                   (pSrc[0 + stride2] * m.BottomLeft) +
                                   (pSrc[3 + stride2] * m.BottomMid) +
                                   (pSrc[6 + stride2] * m.BottomRight))
                                   / m.Factor) + m.Offset;

                        if (nPixel < 0) nPixel = 0;
                        if (nPixel > 255) nPixel = 255;
                        p[3 + stride] = (byte)nPixel;

                        p += 3;
                        pSrc += 3;
                    }

                    p += nOffset;
                    pSrc += nOffset;
                }
            }

            b.UnlockBits(bmData);
            bSrc.UnlockBits(bmSrc);
            return true;
        }
        public static bool Smooth(Bitmap b, int weight = 1)
        {
            ConvMatrix m = new ConvMatrix();
            m.SetAll(1);
            m.Pixel = weight;
            m.Factor = weight + 8;
            return Conv3x3(b, m);
        }

        public static bool GaussianBlur(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopLeft = m.BottomRight = m.TopRight = m.BottomLeft = 1;
            m.TopMid = m.MidLeft = m.MidRight = m.BottomMid = 2;
            m.Pixel = 4;
            m.Factor = 16;
            return Conv3x3(b, m);
        }

        public static bool Sharpen(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopLeft = m.BottomRight = m.TopRight = m.BottomLeft = 0;
            m.TopMid = m.MidLeft = m.MidRight = m.BottomMid = -2;
            m.Pixel = 11;
            m.Factor = 3;
            return Conv3x3(b, m);
        }

        public static bool MeanRemoval(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.SetAll(-1);
            m.Pixel = 9;
            m.Factor = 1;
            return Conv3x3(b, m);
        }

        public static bool EmbossLaplacian(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopLeft = -1; m.TopMid = 0; m.TopRight = -1;
            m.MidLeft = 0; m.Pixel = 4; m.MidRight = 0;
            m.BottomLeft = -1; m.BottomMid = 0; m.BottomRight = -1;
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }

        public static bool EmbossHorzVert(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopMid = -1; m.MidLeft = -1; m.Pixel = 4; m.MidRight = -1; m.BottomMid = -1;
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }

        public static bool EmbossAllDir(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.SetAll(-1);
            m.Pixel = 8;
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }

        public static bool EmbossLossy(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopLeft = 1; m.TopMid = -2; m.TopRight = 1;
            m.MidLeft = -2; m.Pixel = 4; m.MidRight = -2;
            m.BottomLeft = -2; m.BottomMid = 1; m.BottomRight = -2;
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }

        public static bool EmbossHorizontal(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.MidLeft = -1; m.Pixel = 2; m.MidRight = -1;
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }

        public static bool EmbossVertical(Bitmap b)
        {
            ConvMatrix m = new ConvMatrix();
            m.TopMid = -1; m.BottomMid = 1; 
            m.Factor = 1; m.Offset = 127;
            return Conv3x3(b, m);
        }
    }
}