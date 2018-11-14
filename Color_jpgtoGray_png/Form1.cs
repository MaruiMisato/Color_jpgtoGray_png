using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using OpenCvSharp;
using System.Drawing;
using System.Drawing.Imaging;

namespace Color_jpgtoGray_png {
    public partial class Form1:Form {
        public Form1() {
            InitializeComponent();
        }
        private unsafe void GetHistgram(IplImage GrayImage,int[] Histgram){
            byte* g=(byte*)GrayImage.ImageData;
            for(int l=0;l<GrayImage.ImageSize;Histgram[g[l++]]++);
        }
        private void SaveAsPNG(string f,IplImage GrayImage){
            Cv.SaveImage(System.IO.Path.ChangeExtension(f,"png"),GrayImage,new ImageEncodingParam(ImageEncodingID.PngCompression,0));
        }
        private  unsafe void ReducedToBit(IplImage p_img) {//2bit に減色
            byte* p=(byte*)p_img.ImageData;
            for(int y=0;y<p_img.ImageSize;++y) 
                p[y]=(byte)(((p[y]>>4)<<4)+(p[y]>>4));
        }
        private  unsafe void ReducedTo4Bit(IplImage p_img) {//4bit に減色
            byte* p=(byte*)p_img.ImageData;
            for(int y=0;y<p_img.ImageSize;++y) 
                p[y]=(byte)(((p[y]>>6)<<6)+((p[y]>>6)<<4)+((p[y]>>6)<<2)+((p[y]>>6)));
        }
        private bool YellowStainRemovalEntry(string f,TextWriter writerSync){
            IplImage GrayImage=Cv.LoadImage(f,LoadMode.GrayScale);//gray
            int[] Histgram=new int[Const.Tone8Bit];
            GetHistgram(GrayImage,Histgram);
            Image.ToneValue ImageToneValue =new Image.ToneValue();
            ImageToneValue.Max=Image.GetToneValueMax(Histgram);
            ImageToneValue.Min=Image.GetToneValueMin(Histgram);//(byte)がないと豆腐でエラー
            writerSync.WriteLine(f+"\nToneValue.Max:"+ImageToneValue.Max.ToString()+"//ToneValue.Min:"+ImageToneValue.Min.ToString());
            if(ImageToneValue.Max<=ImageToneValue.Min){
                Cv.ReleaseImage(GrayImage);//豆腐･アスファルトはスルー
                return false;
            }
            Image.Transform2Linear(GrayImage,ImageToneValue);
            if(radioButton1.Checked)//4bit
                ReducedToBit(GrayImage);
            else if(radioButton3.Checked)//2bit
                ReducedTo4Bit(GrayImage);
            SaveAsPNG(f,GrayImage);
            Cv.ReleaseImage(GrayImage);
            System.IO.File.Delete(f);
            return true;
         }
        private void button1_Click(object sender,EventArgs e) {
            if(!Clipboard.ContainsFileDropList()) {//Check if clipboard has file drop format data. 取得できなかったときはnull listBox1.Items.Clear();
                MessageBox.Show("Please select files.");
            }
            string[] AllFile=new string[Clipboard.GetFileDropList().Count];
            Clipboard.GetFileDropList().CopyTo(AllFile,0);
            using(TextWriter writerSync=TextWriter.Synchronized(new StreamWriter(DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")+".txt",false,System.Text.Encoding.GetEncoding("shift_jis")))) { 
                Parallel.ForEach(AllFile,new ParallelOptions() { MaxDegreeOfParallelism=4 },f => {
                    YellowStainRemovalEntry(f,writerSync);
                });
                writerSync.WriteLine(DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss"));
            }
        }
    }
}
