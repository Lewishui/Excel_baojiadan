using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ThoughtWorks.QRCode
{
    public partial class frmtest : Form
    {
        public frmtest()
        {
            InitializeComponent();
        }
        private void scanbitmap()
        {
            ////if (TextUtils.isEmpty(path))
            ////{

            ////    return null;

            ////}
            //// DecodeHintType 和EncodeHintType
            //Hashtable<DecodeHintType, String> hints = new Hashtable<DecodeHintType, String>();
            //hints.put(DecodeHintType.CHARACTER_SET, "utf-8"); // 设置二维码内容的编码
            //BitmapFactory.Options options = new BitmapFactory.Options();
            //options.inJustDecodeBounds = true; // 先获取原大小
            //scanBitmap = BitmapFactory.decodeFile(path, options);
            //options.inJustDecodeBounds = false; // 获取新的大小

            //int sampleSize = (int)(options.outHeight / (float)200);

            //if (sampleSize <= 0)
            //    sampleSize = 1;
            //options.inSampleSize = sampleSize;
            //scanBitmap = BitmapFactory.decodeFile(path, options);

            /////
            //LuminanceSource source1 = new PlanarYUVLuminanceSource(
            //  rgb2YUV(scanBitmap), scanBitmap.getWidth(),
            //  scanBitmap.getHeight(), 0, 0, scanBitmap.getWidth(),
            //  scanBitmap.getHeight());
            //BinaryBitmap binaryBitmap = new BinaryBitmap(new HybridBinarizer(
            //        source1));
            //MultiFormatReader reader1 = new MultiFormatReader();
            //Result result1;
            //try
            //{
            //    result1 = reader1.decode(binaryBitmap);
            //    String content = result1.getText();
            //    Log.e("123content", content);
            //}
            //catch (NotFoundException e1)
            //{
            //    e1.printStackTrace();
            //}


        }
  

        //  private String recode(String str)
        //{
        //    String format = "";

        //    try
        //    {
        //        Boolean ISO = Charset.forName("ISO-8859-1").newEncoder()
        //                .canEncode(str);
        //        if (ISO)
        //        {
        //            format = new String(str.("ISO-8859-1"), "GB2312");
        //          //  Log.i("1234      ISO8859-1", format);
        //        }
        //        else
        //        {
        //            format = str;
        //           // Log.i("1234      stringExtra", str);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        e.printStackTrace();
        //    }
        //    return format;
        //}
        public byte[] rgb2YUV(Bitmap bitmap)
        {
            int width = bitmap.Width;
            int height = bitmap.Height;
            int[] pixels = new int[width * height];
            bitmap.GetPixel(width, height);

            int len = width * height;
            byte[] yuv = new byte[len * 3 / 2];
            int y, u, v;
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    int rgb = pixels[i * width + j] & 0x00FFFFFF;

                    int r = rgb & 0xFF;
                    int g = (rgb >> 8) & 0xFF;
                    int b = (rgb >> 16) & 0xFF;

                    y = ((66 * r + 129 * g + 25 * b + 128) >> 8) + 16;
                    u = ((-38 * r - 74 * g + 112 * b + 128) >> 8) + 128;
                    v = ((112 * r - 94 * g - 18 * b + 128) >> 8) + 128;

                    y = y < 16 ? 16 : (y > 255 ? 255 : y);
                    u = u < 0 ? 0 : (u > 255 ? 255 : u);
                    v = v < 0 ? 0 : (v > 255 ? 255 : v);

                    yuv[i * width + j] = (byte)y;
                    //                yuv[len + (i >> 1) * width + (j & ~1) + 0] = (byte) u;
                    //                yuv[len + (i >> 1) * width + (j & ~1) + 1] = (byte) v;
                }
            }
            return yuv;
        }
    
    }
}
