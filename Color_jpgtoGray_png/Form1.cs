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
        private void NoiseRemoveTwoArea(IplImage p_img,byte max) {
            using(IplImage q_img=Cv.CreateImage(Cv.GetSize(p_img),BitDepth.U8,1)) {
                unsafe {
                    byte* p=(byte*)p_img.ImageData,q=(byte*)q_img.ImageData;
                    for(int y=0;y<q_img.ImageSize;++y)q[y]=p[y]<max?(byte)0:(byte)255;//First, binarize
                    for(int y=1;y<q_img.Height-1;y++) {
                        int yoffset=(q_img.WidthStep*y);
                        for(int x=1;x<q_img.Width-1;x++) {
                            int offset=yoffset+x;
                            if(q[offset]==0)//Count white spots around black dots
                                for(int yy=-1;yy<2;++yy) {
                                    int yyyoffset=q_img.WidthStep*(y+yy);
                                    for(int xx=-1;xx<2;++xx) if(q[yyyoffset+(x+xx)]==255)++q[offset];//自身は数えられないから8が最高
                                }
                        }
                    }
                    for(int y=1;y<q_img.Height-1;y++) {
                        int yoffset=(q_img.WidthStep*y);
                        for(int x=1;x<q_img.Width-1;x++) {
                            if(q[yoffset+x]==7)//When there are seven white spots in the periphery
                                for(int yy=-1;yy<2;++yy) {
                                    int yyyoffset = q_img.WidthStep*(y+yy);
                                    for(int xx=-1;xx<2;++xx) {
                                        int offset=yyyoffset+(x+xx);
                                        if(q[offset]==7) {//仲間 ペア
                                            p[yoffset+x]=max;//q[yoffset+p_img.NChannels*x]=6;//Unnecessary 
                                            p[offset]=max;q[offset]=6;
                                            yy=1;break;
                                        } else;
                                    }
                                }
                            else if(q[yoffset+x]==8)p[yoffset+x]=max;//Independent
                        }
                    }
                }
            }
        }
        private void HiFuMiYoWhite(IplImage p_img,byte threshold,ref int hi,ref int fu,ref int mi,ref int yo) {
            unsafe {
                byte* p=(byte*)p_img.ImageData;
                for(int y=0;y<p_img.Height-1;y++) {//Y上取得
                    int l=0;
                    for(int yy=0;((l==yy)&&(yy<5)&&(y+yy)<p_img.Height-1);++yy) {
                        int yyyoffset=(p_img.WidthStep*(y+yy));
                        for(int x=0;x<p_img.Width-1;x++) if(p[yyyoffset+x]<threshold) { l=yy+1; break; }
                    }if(l==5) {hi=y;break;} 
                    else y+=l;
                }for(int y=p_img.Height-1;y>hi;--y) {//Y下取得
                    int l=0;
                    for(int yy=0;((l==yy)&&(yy>-5)&&(y+yy)>hi);--yy) {
                        int yyyoffset=(p_img.WidthStep*(y+yy));
                        for(int x=0;x<p_img.Width-1;x++)if(p[yyyoffset+x]<threshold) { l=yy-1; break; }
                    }if(l==-5) {mi=y;break;}
                    else y+=l;
                }for(int x=0;x<p_img.Width-1;x++) {//X左取得
                    int l=0;
                    for(int xx=0;((l==xx)&&(xx<5)&&(x+xx)<p_img.Width-1);++xx) {
                        int xxxoffset=(x+xx);
                        for(int y=hi;y<mi-1;y++) if(p[p_img.WidthStep*y+xxxoffset]<threshold) { l=xx+1; break; }
                    }if(l==5) { fu=x;break; }
                    else x+=l;
                }for(int x=p_img.Width-1;x>fu;--x) {//X右取得
                    int l=0;
                    for(int xx=0;((l==xx)&&(xx>-5)&&(x+xx)>fu);--xx) {
                        int xxxoffset=(x+xx);
                        for(int y=hi;y<mi;y++)if(p[p_img.WidthStep*y+xxxoffset]<threshold) { l=xx-1; break; }
                    }if(l==-5) { yo=x;break; }
                    else x+=l;
                }
            }
        }
        private void NoiseRemoveWhite(IplImage p_img,byte min) {//White spot will be remove
            using(IplImage q_img=Cv.CreateImage(Cv.GetSize(p_img),BitDepth.U8,1)) {
                unsafe {
                    byte* p=(byte*)p_img.ImageData,q=(byte*)q_img.ImageData;
                    for(int y=0;y<q_img.ImageSize;++y)q[y]=p[y]>min?(byte)255:(byte)0;//First, binarize
                    for(int y=1;y<q_img.Height-1;++y) {
                        int yoffset=(q_img.WidthStep*y);
                        for(int x=1;x<q_img.Width-1;++x) {
                            int offset=yoffset+x;
                            if(q[offset]==255)//Count white spots around black dots
                                for(int yy=-1;yy<2;++yy) {
                                    int yyyoffset=q_img.WidthStep*(y+yy);
                                    for(int xx=-1;xx<2;++xx) if(q[yyyoffset+(x+xx)]==0)--q[offset];
                                }
                        }
                    }
                    for(int y=1;y<q_img.Height-1;++y) {
                        int yoffset=(q_img.WidthStep*y);
                        for(int x=1;x<q_img.Width-1;++x) {
                            if(q[yoffset+x]==248)//When there are seven white spots in the periphery
                                for(int yy=-1;yy<2;++yy) {
                                    int yyyoffset = q_img.WidthStep*(y+yy);
                                    for(int xx=-1;xx<2;++xx) {
                                        int offset=yyyoffset+(x+xx);
                                        if(q[offset]==248) {//仲間 ペア
                                            p[yoffset+x]=min;//q[offset]=0;//Unnecessary 
                                            p[offset]=min;q[offset]=0;
                                            yy=1;break;
                                        } else;
                                    }
                                }
                            else if(q[yoffset+x]==247)p[yoffset+x]=min;//Independent
                        }
                    }
                }
            }
        }        
        private void HiFuMiYoBlack(IplImage p_img,byte threshold,ref int hi,ref int fu,ref int mi,ref int yo) {
            unsafe {
                byte* p=(byte*)p_img.ImageData;
                for(int y=0;y<p_img.Height-1-4;y++) {//Y上取得
                    if(y==0){
                        int[] l=new int[5];
                        for(int yy = 0;((yy<5)&&(y+yy)<p_img.Height-1);++yy) {
                            int yyyoffset = (p_img.WidthStep*(y+yy));
                            for(int x = 0;x<p_img.Width;x++)if(p[yyyoffset+x]<threshold)++l[yy];
                        }
                        if((p_img.Width>l[0])&&(p_img.Width>l[1])&&(p_img.Width>l[2])&&(p_img.Width>l[3])&&(p_img.Width>l[4])) { hi=y; break; }
                    } else {
                        int l=0;
                        int yoffset=(p_img.WidthStep*(y+4));
                        for(int x = 0;x<p_img.Width;x++)if(p[yoffset+x]<threshold) ++l;
                        if((p_img.Width>l)) { hi=y; break; }
                    }
                }for(int y=p_img.Height-1;y>(hi+4);--y) {//Y下取得
                    if(y==p_img.Height-1) {
                        int[] l=new int[5];
                        for(int yy=-4;((yy<1)&&(y+yy)>hi);++yy) {
                            int yyyoffset=(p_img.WidthStep*(y+yy));
                            for(int x=0;x<p_img.Width;x++) if(p[yyyoffset+x]<threshold)++l[-yy];
                        }
                        if((p_img.Width>l[0])&&(p_img.Width>l[1])&&(p_img.Width>l[2])&&(p_img.Width>l[3])&&(p_img.Width>l[4])) { mi=y; break; }
                    } else {
                        int yoffset=(p_img.WidthStep*(y-4));
                        int l=0;
                        for(int x=0;x<p_img.Width;x++) if(p[yoffset+x]<threshold)++l;
                        if((p_img.Width>l)) { mi=y; break; }
                    }
                }for(int x=0;x<p_img.Width-1-4;x++) {//X左取得
                    if(x==0) {
                        int[] l=new int[5];
                        for(int xx=0;((xx<5)&&(x+xx)<p_img.Width-1);++xx) {
                            int xxxoffset=(x+xx);
                            for(int y=hi;y<mi;y++) if(p[xxxoffset+p_img.WidthStep*y]<threshold)++l[xx];
                        }
                        if(((mi-hi)>l[0])&&((mi-hi)>l[1])&&((mi-hi)>l[2])&&((mi-hi)>l[3])&&((mi-hi)>l[4])) { fu=x; break; }
                    } else {
                        int xoffset=(x+4);
                        int l=0;
                        for(int y=hi;y<mi;y++) if(p[xoffset+p_img.WidthStep*y]<threshold)++l;
                        if((mi-hi)>l) { fu=x; break; }
                    }
                }for(int x=p_img.Width-1;x>(fu+4);--x) {//X右取得
                    if(x==p_img.Width-1) {
                        int[] l=new int[5];
                        for(int xx=-4;((xx<0)&&(x+xx)>fu);++xx) {
                            int xxxoffset=(x+xx);
                            for(int y=hi;y<mi;y++) if(p[xxxoffset+p_img.WidthStep*y]<threshold)++l[-xx];
                        }
                        if(((mi-hi)>l[0])&&((mi-hi)>l[1])&&((mi-hi)>l[2])&&((mi-hi)>l[3])&&((mi-hi)>l[4])) { yo=x; break; }
                    } else {
                        int xoffset=(x-4);
                        int l=0;
                        for(int y=hi;y<mi;y++) if(p[xoffset+p_img.WidthStep*y]<threshold)++l;
                        if((mi-hi)>l) { yo=x; break; }
                    }
                }
            }
        }
        private void WhiteCut(IplImage p_img,IplImage q_img,int hi,int fu,int mi,int yo) {
            unsafe {
                byte* p=(byte*)p_img.ImageData,q=(byte*)q_img.ImageData;
                for(int y=hi;y<=mi;++y) {
                    int yoffset=(p_img.WidthStep*y),qyoffset=(q_img.WidthStep*(y-hi));
                    for(int x=fu;x<=yo;++x)q[qyoffset+(x-fu)]=p[yoffset+x];
                }
            }
        }
        private void GetSpaces(IplImage p_img,byte threshold,int[] ly,int[] lx,ref int new_h,ref int new_w) {
            unsafe {
                byte* p=(byte*)p_img.ImageData;
                for(int y=1;y<p_img.Height;++y) {
                    int yoffset=p_img.WidthStep*y;
                    int ybefore=ly[y-1]+1;
                    int blackstone=0;
                    for(int x=0;x<p_img.Width;++x)
                        if(p[yoffset+x]>threshold)ly[y]=ybefore;//白
                        else if(++blackstone>(2)){ ly[y]=0; break; }//黒2つ以上
                }
                int w_width=(int)((p_img.Height+p_img.Width)*0.02);//(p_img.Height+p_img.Width)*0.02)は残す空白の大きさ
                for(int y=1;y<p_img.Height;++y) if(ly[y]<=w_width) { ly[y]=0; ++new_h; }//(p_img.Height+p_img.Width)*0.02)は残す空白の大きさ
                for(int x=1;x<p_img.Width;++x) {
                    int xbefore=lx[x-1]+1;
                    int blackstone=0;
                    for(int y=0;y<p_img.Height;++y)
                        if(p[p_img.WidthStep*y+x]>threshold)lx[x]=xbefore;//白
                        else if(++blackstone>(2)){ lx[x]=0; break; }//黒2つ以上
                }
                for(int x=1;x<p_img.Width;++x) if(lx[x]<=w_width) { lx[x]=0; ++new_w; }//(p_img.Height+p_img.Width)*0.02)は残す空白の大きさ
            }
        }
        private void DeleteSpaces(ref string f,IplImage p_img,byte threshold,byte min,double magnification) {//内部の空白を除去 グレイスケールのみ
            int[] ly=new int[p_img.Height],lx=new int[p_img.Width];
            int new_h=+1,new_w=+1;
            GetSpaces(p_img,threshold,ly,lx,ref new_h,ref new_w);
            using(IplImage q_img=Cv.CreateImage(new CvSize(new_w,new_h),BitDepth.U8,1)) {
                unsafe {
                    byte* q=(byte*)q_img.ImageData,p=(byte*)p_img.ImageData;
                    if(radioButton1.Checked==true)
                        for(int y=0,yy=0;y<p_img.Height;++y)
                            if(ly[y]==0) {
                                int yyoffset=q_img.WidthStep*(yy++),yoffset=p_img.WidthStep*y;
                                for(int x=0,xx=0;x<p_img.Width;++x) 
                                    if(lx[x]==0) {
                                    int offset=yyoffset+(xx++);
                                        q[offset]=(byte)(magnification*(p[yoffset+x]-min));//255.99ないと255が254になる
                                        q[offset]=(byte)(((q[offset]>>4)<<4)+(q[offset]>>4));
                                    }
                            } else ;
                    else
                        for(int y=0,yy=0;y<p_img.Height;++y)
                            if(ly[y]==0) {
                                int yyoffset=q_img.WidthStep*(yy++),yoffset=p_img.WidthStep*y;
                                for(int x=0,xx=0;x<p_img.Width;++x) if(lx[x]==0) q[yyoffset+(xx++)]=(byte)(magnification*(p[yoffset+x]-min));//255.99ないと255が254になる
                            } else ;
                }
                Cv.SaveImage(f,q_img,new ImageEncodingParam(ImageEncodingID.PngCompression,0));
            }
        }
        private void DeleteSpaces(ref string f,IplImage p_img,byte threshold) {//内部の空白を除去 グレイスケールのみ
            int[] ly=new int[p_img.Height],lx=new int[p_img.Width];
            int new_h=+1,new_w=+1;
            GetSpaces(p_img,threshold,ly,lx,ref new_h,ref new_w);
            using(IplImage q_img=Cv.CreateImage(new CvSize(new_w,new_h),BitDepth.U8,1)) {
                unsafe {
                    byte* q=(byte*)q_img.ImageData,p=(byte*)p_img.ImageData;
                    if(radioButton1.Checked==true)
                        for(int y=0,yy=0;y<p_img.Height;++y)
                            if(ly[y]==0) {
                                int yyoffset=q_img.WidthStep*(yy++),yoffset=p_img.WidthStep*y;
                                for(int x=0,xx=0;x<p_img.Width;++x) if(lx[x]==0)q[yyoffset+(xx++)]=(byte)(((p[yoffset+x]>>4)<<4)+(p[yoffset+x]>>4));//4bitに減色 下位bitを同じので埋める
                            }else;
                    else {
                        for(int y=0,yy=0;y<p_img.Height;++y)
                        if(ly[y]==0) {
                            int yyoffset=q_img.WidthStep*(yy++),yoffset=p_img.WidthStep*y;
                            for(int x=0,xx=0;x<p_img.Width;++x) if(lx[x]==0)q[yyoffset+(xx++)]=p[yoffset+x];//8bit保存
                        }else;
                    }
                }
                Cv.SaveImage(f,q_img,new ImageEncodingParam(ImageEncodingID.PngCompression,0));
            }
        }
        private void PNGRemoveAlways(string f,uint n) {
            while(0!=n--) 
                using(IplImage g_img=Cv.LoadImage(f,LoadMode.GrayScale)) {
                    int[] histgram=new int[256];
                    unsafe {
                        byte* g=(byte*)g_img.ImageData;
                        for(int l=0;l<g_img.ImageSize;histgram[g[l++]]++);
                    }
                    byte i=255;
                    for(int total=0;(total+=histgram[i])<g_img.ImageSize*0.6;--i);byte threshold=--i;//0.1~0.7/
                    NoiseRemoveTwoArea(g_img,255);
                    NoiseRemoveWhite(g_img,0);
                    int hi=0,fu=0,mi=g_img.Height-1,yo=g_img.Width-1;
                    HiFuMiYoWhite(g_img,threshold,ref hi,ref fu,ref mi,ref yo);
                    if((hi==0)&&(fu==0)&&(mi==g_img.Height-1)&&(yo==g_img.Width-1))HiFuMiYoBlack(g_img,127,ref hi,ref fu,ref mi,ref yo);//background black
                    using(IplImage p_img=Cv.CreateImage(new CvSize((yo-fu)+1,(mi-hi)+1),BitDepth.U8,1)) {
                        WhiteCut(g_img,p_img,hi,fu,mi,yo);
                        DeleteSpaces(ref f,p_img,threshold);//内部の空白を除去 階調値変換
                    } 
                }
        }
        private void PNGOut(string[] all_file) {
            System.Diagnostics.Process p=new System.Diagnostics.Process();//Create a Process object
            p.StartInfo.FileName=System.Environment.GetEnvironmentVariable("ComSpec");//ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.WindowStyle=System.Diagnostics.ProcessWindowStyle.Hidden;//HiddenMaximizedMinimizedNormal
            Parallel.ForEach(all_file,new ParallelOptions() { MaxDegreeOfParallelism=4 },f => {
                p.StartInfo.Arguments="/c pngout "+System.IO.Path.ChangeExtension(f,"png");//By default, PNGOUT will not overwrite a PNG file if it was not able to compress it further.
                //p.StartInfo.Arguments="/c pngout \""+System.IO.Path.ChangeExtension(f,"png")+"\"";
                p.Start();p.WaitForExit();//起動
            });
            p.Close();
        }
        private void button1_Click(object sender,EventArgs e) {
            if(Clipboard.ContainsFileDropList()) {//Check if clipboard has file drop format data. 取得できなかったときはnull listBox1.Items.Clear();
                string[] all_file=new string[Clipboard.GetFileDropList().Count];
                int j=0;
                foreach(string f in Clipboard.GetFileDropList()) all_file[j++]=f;
                using(TextWriter writerSync=TextWriter.Synchronized(new StreamWriter(DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")+".txt",false,System.Text.Encoding.GetEncoding("shift_jis")))) { 
                    Parallel.ForEach(all_file,new ParallelOptions() { MaxDegreeOfParallelism=4 },f => {
                        string f2=System.IO.Path.ChangeExtension(f,"png");
                        using(IplImage g_img=Cv.LoadImage(f,LoadMode.GrayScale)) {//gray
                            int[] histgram=new int[256];
                            unsafe {
                                byte* g=(byte*)g_img.ImageData;
                                for(int l=0;l<g_img.ImageSize;histgram[g[l++]]++);
                            }
                            byte i=256-2;
                            for(/*i=256-2*/;histgram[(byte)(i+1)]==0;--i);byte max=++i;//(byte)がないとアスファルトでエラー
                            for(i=1;histgram[(byte)(i-1)]==0;++i);byte min=--i;//(byte)がないと豆腐でエラー
                            if(max>min) {//豆腐･アスファルトはスルー
                                i=255;
                                for(int total=0;(total+=histgram[i])<g_img.ImageSize*0.6;--i);byte threshold=--i;//0.1~0.7/p_img.NChannels
                                NoiseRemoveTwoArea(g_img,max);//colorには後ほど反映させる
                                NoiseRemoveWhite(g_img,min);
                                int hi=0,fu=0,mi=g_img.Height-1,yo=g_img.Width-1;
                                HiFuMiYoWhite(g_img,threshold,ref hi,ref fu,ref mi,ref yo);
                                if((hi==0)&&(fu==0)&&(mi==g_img.Height-1)&&(yo==g_img.Width-1))HiFuMiYoBlack(g_img,(byte)((max-min)>>1),ref hi,ref fu,ref mi,ref yo);//background black
                                using(IplImage p_img=Cv.CreateImage(new CvSize((yo-fu)+1,(mi-hi)+1),BitDepth.U8,1)) {
                                    writerSync.WriteLine(f+"\n\tth="+threshold+":min="+min+":max="+max+":hi="+hi+":fu="+fu+":mi="+mi+":yo="+yo+"\n\tHeight="+g_img.Height+":Width="+g_img.Width+"\n\theight="+p_img.Height+":width="+p_img.Width);
                                    WhiteCut(g_img,p_img,hi,fu,mi,yo);
                                    DeleteSpaces(ref f2,p_img,(byte)(threshold*(255.99/(max-min))),min,255.99/(max-min));//内部の空白を除去 階調値変換
                                }
                                PNGRemoveAlways(f2,4);//n回繰り返す
                                System.IO.File.Delete(System.IO.Path.ChangeExtension(f,"jpg"));//Disposal of garbage//System.IO.File.Move(f,System.IO.Path.ChangeExtension(f,"png"));//実際にファイル名を変更する
                            }
                        }
                    });
                    writerSync.WriteLine(DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss"));
                }
                PNGOut(all_file);
            } else
                MessageBox.Show("Please select files.");
        }
    }
}
