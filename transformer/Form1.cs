using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.License;
using Spire.Pdf;
using Spire.Doc;
using System.IO;
using Spire.Pdf.Exporting.XPS.Schema;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Reflection;

namespace transformer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            InitializeComponent();


        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }
       System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string dllName = args.Name.Contains(",") ? args.Name.Substring(0, args.Name.IndexOf(',')) : args.Name.Replace(".dll", "");
            dllName = dllName.Replace(".", "_");
            if (dllName.EndsWith("_resources")) return null;
            System.Resources.ResourceManager rm = new System.Resources.ResourceManager(GetType().Namespace + ".Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());
            byte[] bytes = (byte[])rm.GetObject(dllName);
            return System.Reflection.Assembly.Load(bytes);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            #region   将pdf分成许多份小文档
            Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            pdf.LoadFromFile(textBox1.Text);
            label4.Text = "转换中......";
            label4.Refresh();
            for(int i=0;i < pdf.Pages.Count;i += 5)
            {
                
                int j=0;
                Spire.Pdf.PdfDocument newpdf = new Spire.Pdf.PdfDocument();
                for(j=i;j>=i&&j<=i+4;j++)
                {
                    if(j<pdf.Pages.Count)
                    {
                        Spire.Pdf.PdfPageBase page;
                        page = newpdf.Pages.Add(pdf.Pages[j].Size, new Spire.Pdf.Graphics.PdfMargins(0));
                        pdf.Pages[j].CreateTemplate().Draw(page,new PointF(0,0));
                    }
                    
                }
                newpdf.SaveToFile(textBox2.Text+"\\"+j.ToString()+".pdf");
                PdfExtractWordAndPicture(textBox2.Text,j.ToString());
            }
            #endregion


            #region  合并word文档

            string filePath0 = textBox2.Text + "\\" +'5' + ".doc";
            for(int i=10;i<=0-pdf.Pages.Count%5+pdf.Pages.Count;i+=5)
            {
                string filePath2 = textBox2.Text + "\\" + i.ToString() + ".doc";

                Spire.Doc.Document doc = new Spire.Doc.Document(filePath0);
                doc.InsertTextFromFile(filePath2,Spire.Doc.FileFormat.Doc);

                doc.SaveToFile(filePath0,Spire.Doc.FileFormat.Doc);
            }
            Spire.Doc.Document mydoc1 = new Spire.Doc.Document();
            mydoc1.LoadFromFile(textBox2.Text + "\\" + '5' + ".doc");
            mydoc1.SaveToFile(textBox2.Text + "\\" + "TheLastTransform" + ".doc", Spire.Doc.FileFormat.Doc);

            for (int i = 5; i <= 5 - pdf.Pages.Count % 5 + pdf.Pages.Count; i += 5)
            {
                File.Delete(textBox2.Text + "\\" + i.ToString() + ".doc");
                File.Delete(textBox2.Text + "\\" + i.ToString() + ".pdf");
            }
            
            #endregion

            label4.Text = "转换完成";
            label4.Refresh();
        }
        private void PdfExtractWordAndPicture(string savePathCache,string midName)//参数：保存地址，处理过程中文件名称(不包含后缀)
        { 
            #region   提取PDF中的文字
            
            try
            {   
                    PdfDocument doc = new PdfDocument();
                    doc.LoadFromFile(savePathCache+"\\"+midName+".pdf");        //加载文件
                    StringBuilder content = new StringBuilder();
                    foreach (PdfPageBase page in doc.Pages)
                    {
                        content.Append(page.ExtractText());
                    }
                  

                    System.IO.File.WriteAllText(savePathCache+"\\mid.txt", content.ToString());

                    Spire.Doc.Document document = new Spire.Doc.Document();
                    document.LoadFromFile(savePathCache + "\\mid.txt");
                    document.Replace(" ", "", true, true);
                    document.Replace("Evaluation Warning : The document was created with Spire.PDF for .NET.", "", true, true);
                    document.SaveToFile(savePathCache +"\\"+ midName+".doc", Spire.Doc.FileFormat.Doc);

                    File.Delete(savePathCache + "\\mid.txt");

            }
            catch (Exception)
            { 
                MessageBox.Show("请填写正确的路径");
            }
            #endregion

            #region  提取PDF中的图片
           //创建一个PdfDocument类对象并加载PDF sample
            Spire.Pdf.PdfDocument mydoc = new Spire.Pdf.PdfDocument();
            mydoc.LoadFromFile(savePathCache + "\\" + midName + ".pdf"); 

            //声明一个IList类，元素为image
            IList<Image> images = new List<Image>();
            //遍历PDF文档中诊断是否包含图片，并提取图片
            foreach (PdfPageBase page in mydoc.Pages)
            {
            if (page.ExtractImages() != null)
               {
                 foreach (Image image in page.ExtractImages())
                     {
                           images.Add(image);
                      }
               }
            }
            mydoc.Close();

            //遍历提取的图片，保存并命名图片
            int index = 0;
            foreach (Image image in images)
            {
              String imageFileName = String.Format(midName+"Image-{0}.png", index++);
              image.Save(savePathCache+"\\"+imageFileName, ImageFormat.Png);
             }
            #endregion
        }


        private void PdfToWordSimple(string savePathCache, string midName)
        {
            Spire.Pdf.PdfDocument mydoc2 = new PdfDocument();
            mydoc2.LoadFromFile(savePathCache + "\\" + midName + ".pdf");
            string a = savePathCache + "\\" + midName + ".doc";
            mydoc2.SaveToFile(a,Spire.Pdf.FileFormat.DOCX);
        }
        private void button3_Click(object sender, EventArgs e)                  //打开转换文件路径
        {
            string file="";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.*)|*.*";        
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                file = dialog.FileName;
            }
            textBox1.Text = file;
        }

        private void button4_Click(object sender, EventArgs e)                  //保存的路径
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                DirectoryInfo theFolderSave = new DirectoryInfo(foldPath);

                //theFolder 包含文件路径
                textBox2.Text = theFolderSave.ToString();
            }
        }

        //private void button5_Click(object sender, EventArgs e)                    //打开文件
        //{
        //    try
        //    {
        //        System.Diagnostics.Process.Start(savePath);
        //    }
        //    catch (Exception)
        //    {
        //        MessageBox.Show("请正确操作");
        //    }
            
        //}

        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{
            //#region   将pdf分成许多份小文档
            //Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            //pdf.LoadFromFile(textBox1.Text);
            //label4.Text = "转换中......";
            //label4.Refresh();
            //for (int i = 0; i < pdf.Pages.Count; i += 5)
            //{

            //    int j = 0;
            //    Spire.Pdf.PdfDocument newpdf = new Spire.Pdf.PdfDocument();
            //    for (j = i; j >= i && j <= i + 4; j++)
            //    {
            //        if (j < pdf.Pages.Count)
            //        {
            //            Spire.Pdf.PdfPageBase page;
            //            page = newpdf.Pages.Add(pdf.Pages[j].Size, new Spire.Pdf.Graphics.PdfMargins(0));
            //            pdf.Pages[j].CreateTemplate().Draw(page, new PointF(0, 0));
            //        }

            //    }
            //    newpdf.SaveToFile(textBox2.Text + "\\" + j.ToString() + ".pdf");
            //    newpdf.Close();
            //}
            //#endregion

            #region  PDF转word
            //for (int i = 5; i <= ((pdf.Pages.Count%5==0)?pdf.Pages.Count:(pdf.Pages.Count+pdf.Pages.Count%5+5)); i += 5)
            //{
                //PdfToWordSimple(textBox2.Text,i.ToString());
                PdfToWordSimple(textBox2.Text,"10");
                //PdfToWordSimple(textBox2.Text, "20");
            //}

            #endregion
            //#region  合并word文档

                //string filePath0 = textBox2.Text + "\\" + '5' + ".doc";
                //for (int i = 10; i <= 0 - pdf.Pages.Count % 5 + pdf.Pages.Count; i += 5)
                //{
                //    string filePath2 = textBox2.Text + "\\" + i.ToString() + ".doc";

                //    Spire.Doc.Document doc = new Spire.Doc.Document(filePath0);
                //    doc.InsertTextFromFile(filePath2, Spire.Doc.FileFormat.Doc);

                //    doc.SaveToFile(filePath0, Spire.Doc.FileFormat.Doc);
                //}
                //Spire.Doc.Document mydoc1 = new Spire.Doc.Document();
                //mydoc1.LoadFromFile(textBox2.Text + "\\" + '5' + ".doc");
                //mydoc1.SaveToFile(textBox2.Text + "\\" + "TheLastTransform" + ".doc", Spire.Doc.FileFormat.Doc);

                //for (int i = 5; i <= 5 - pdf.Pages.Count % 5 + pdf.Pages.Count; i += 5)
                //{
                //    File.Delete(textBox2.Text + "\\" + i.ToString() + ".doc");
                //    File.Delete(textBox2.Text + "\\" + i.ToString() + ".pdf");
                //}

                //#endregion

                label4.Text = "转换完成";
                label4.Refresh();
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("错误");
            //}
            
        }


    }
         


}
