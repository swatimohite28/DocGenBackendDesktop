using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Xml.Linq;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Paragraph = Xceed.Document.NET.Paragraph;
using Image = Xceed.Document.NET.Image;
using System.Drawing.Imaging;
using CoreHtmlToImage;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace export_html_to_word
{  
   
    public partial class Form1 : Form
    {
        private string htmlFilePath = "";
        private string fileNameWithoutExt = "";
        private string base64code="";
        private string headerImg;
        private string footerText;

        public Form1()
        {
            InitializeComponent();
            //this.BackColor = Color.FromArgb(255, 232, 232);
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
            //string imagePath = @"D:\DocLogo.png";
            //string imgBase64String = GetBase64StringForImage(imagePath);
            //MessageBox.Show(imgBase64String);

            //using (StreamReader r = new StreamReader(System.Windows.Forms.Application.StartupPath + @"../../../Resources/DocGen.json"))
            //{
            //    var json = r.ReadToEnd();
            //    var items = JsonConvert.DeserializeObject<List<SchemaInfo>>(json);
            //    foreach (var item in items)
            //    {
            //        // Console.WriteLine("{0} {1}", item.temp, item.vcc);
            //    }
            //};
            //myMenuStrip.BackColor = Color.LightGreen;

            this.btnOpenFile.BackColor = Color.LightGreen;
            this.btnExport.BackColor = Color.LightGreen;
            string result = string.Empty;

            //using (StreamReader r = new StreamReader(@"../../Resources/DocGen.json"))
            //{
            //    var json = r.ReadToEnd();
            //    var jobj = JObject.Parse(json);

            //    //foreach (var item in jobj.Properties())
            //    //{
            //    //    item.Value = item.Value.ToString().Replace("v1", "v2");
            //    //}
            //    result = jobj.ToString();

            //    //Console.WriteLine(result);
            //   // Console.WriteLine(header);
            //    Console.WriteLine(jobj.GetValue("sectionContent"));
            //    base64code = jobj.GetValue("sectionContent").ToString();
            //    headerImg = jobj.GetValue("SectionName").ToString(); 
            //}


            using (FileStream fs = new FileStream(@"../../Resources/DocGen.json", FileMode.Open, FileAccess.Read))
            using (StreamReader sr = new StreamReader(fs))
            using (JsonTextReader reader = new JsonTextReader(sr))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.StartObject)
                    {
                        // Load each object from the stream
                        JObject obj = JObject.Load(reader);
                        if(obj["sectionType"].ToString() == "Header")
                        {
                            base64code = obj["sectionContent"].ToString();
                            headerImg = obj["SectionName"].ToString();
                        }
                        else if (obj["sectionType"].ToString() == "Footer")
                        {
                            footerText = obj["sectionContent"].ToString();
                        }
                         Console.WriteLine(obj["sectionContent"] + " \n" + obj["SectionName"]);
                    }
                }
            }

        }

        protected static string GetBase64StringForImage(string imgPath)
        {
            byte[] imageBytes = System.IO.File.ReadAllBytes(imgPath);
            string base64String = Convert.ToBase64String(imageBytes);
            return base64String;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = fileDialog.FileName;
                lblFile.Text = Path.GetFileName(sFileName);
                MessageBox.Show(Path.GetFileName(sFileName));
                this.htmlFilePath = sFileName.Replace("\\", "//");
                this.fileNameWithoutExt = Path.GetFileNameWithoutExtension(this.htmlFilePath);
            }

            
        }

        public void ConvertHtmlToImage()
        {
            //Bitmap m_Bitmap = new Bitmap(400, 600);
            //PointF point = new PointF(0, 0);
            //SizeF maxSize = new System.Drawing.SizeF(500, 500);
            //HtmlRenderer.HtmlRender.Render(Graphics.FromImage(m_Bitmap),
            //                                        "<html><body><p>This is some html code</p>"
            //                                        + "<p>This is another html line</p></body>",
            //                                         point, maxSize);

            //m_Bitmap.Save(@"C:\Test.png", ImageFormat.Png);

           // var base64code = "data:image/png;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCACHBAADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+/iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoorl7/wAbeD9L8VaD4G1LxRoNh4y8U2Oral4b8L3eqWcGva7YaEsD6xeaVpcky3l9BpqXML3clvE6wxszsQkcrJM5wppOpOEFKUIRc5xgnOpJQpwTnKKc5yajCCblOTUYRlJ8ptQw+IxMpww2Hr4mdOjXxNSGHoVq86eHw1N1sTiJxoUq0oUMPRTrYivOMaNCknVrVaNNOouooooqjEKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiivj/wDbK/bJ+G37HHw1ufF3i65t9U8W6pb3cPgXwNHdpBqHiLUYEAe7u3Akk03w3pjyQvrOsPC4iV4rKyiu9VvLKyn5Mfj8HlmDxGPx+Ip4XB4WnKrXr1ZKMIQivvlOTtCnTgpTqVJRhCMpSSXucNcNZ9xjn2V8McMZXi85z3OcXSwWW5dgqbqV8RXqtavTko0KMOevicTXlSw2Fw1KtiMRWp0qcpSZ+2Z+2Z8Nv2OPhrceLfFc9vqvjDVbe7i8C+BY7tIL/wAQ38CAPeXjgSSab4a0ySSF9X1doX2+ZDYWEV3qt5Z2c38VXxS/bJ+N3xS+O5+P2seMNWi8aWmr2+paJf6dcS6XJoMenXBm0m18NrE8jaBY6OMrpFnbvIsQeeXU21K71DVJrzkf2kv2kviT+098Stc+I/xH1y41TUNVuP8AR7f54LDTtPgeQ6fpWlaeZJE03R9MjkePTdORnMW+a8vJrvVbu+vrj59r+PePvEXH8WY+MMFOvgclwNdVMBh4zlTq1q1OXuY/FOnJfv7xUqFNSccLFpQbqupWf++H0YPom8M+CPDFbFcRYXL+I/EHiTLZYXiXMq+Hp4rBYDAYulbEcNZPDE0ZpZco1JUswxEqUaucVoyniIRwccLgaf8Abz/wTh/4KOeFf2vvCtn4L8aXmnaH8dtC04tf2CiKxsvH9jYxL9q8Q+H7XdsttYtkxN4m8NQljZFjq+kibQpyNO/U+v8ALX8XftQeKvgp4psLb4NeLNU8N/FPRrqx1pvFOgXZt7jwFLayCfTryKZA6N4mnbIsbN1ZLWzknn1GOSCV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mYSjnvBPwn8T+Ffih4h8YDX9G0zwnqw8Szy+EfC0Piiy03xDq/iLVdP1OHxDr2g6z4l1rwxo3iHSRaXyXmr+DtM0i58YX+uajq2u+UwgsU9/oouK34Hzj4s+DfizXvFniZrHxToVr4A8f+J/h/4v8YafeaHe3Hiuz1L4fN4aEdh4a1SDVbfS00/xHb+ENDtryXVNOuLrRidVubM3z31pFpnS/Fj4L6N8XdS8Iv4hvLiDSPDcHilXi02e90zXodQ1yysINL1zw34k027tNQ8Na74furEXljqtgwvEaRo45Y43mWX2migdjzX4V/Dqz+F3h3UfDGmtYrpUvi3xbr2kWWm2C6bZaPpXiDXbzVdO0SC1SWRNmk2tzFY+cnlrcGEz+TEXKj0qiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/2Q==";
            var converter = new HtmlConverter();
            var html = "<img  align='top' src="+ base64code + " alt='logo' />";
            var bytes = converter.FromHtmlString(html);         
            
            
            if(!File.Exists(@"D:\"+ headerImg + ".png"))
            {
                File.WriteAllBytes(@"D:\" + headerImg + ".png", bytes);
                Console.WriteLine("new image created.");
            }
            else
            {
                Console.WriteLine("image exists.");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ConvertHtmlToImage();

            //Convert the HTML File to Word document.
            //string htmlFilePath = this.fileDialog.FileName;
            _Application word = new Microsoft.Office.Interop.Word.Application();
            _Document wordDoc = word.Documents.Open(FileName: this.htmlFilePath, ReadOnly: false);
            wordDoc.SaveAs(FileName: @"D://" + this.fileNameWithoutExt + ".docx", FileFormat: WdSaveFormat.wdFormatDocumentDefault);
            wordDoc.Close();


            //image
            //string html = "<img src='data:image/png;base64,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' alt='logo' />";
                      

            //Create the HTML file.
            //File.WriteAllText(@"D:\Doclogo.htm", html);
            //MessageBox.Show("HTML File created.");
            


            string fileName = @"D:\"+ this.fileNameWithoutExt + ".docx";
            ////var doc = DocX.Load(fileName);
            ////Create docx  
            //var doc = DocX.Create(fileName);
            ////Create a picture
            //Image img = doc.AddImage(@"D://rose.jpeg");
            //Picture p = img.CreatePicture();
            ////Create a new paragraph  
            //Paragraph par = doc.InsertParagraph("Word picture 2");
            //par.AppendPicture(p);
            //doc.Save();


            // Create a new document.
            using (DocX document = DocX.Load(fileName))
            {
                // Add Header and Footer support to this document.
                document.AddHeaders();
                document.AddFooters();

                // Get the default Header for this document.
                Header header_default = document.Headers.Odd;

                // Get the default Footer for this document.
                Footer footer_default = document.Footers.Odd;

                Image img = document.AddImage(@"D:\" + headerImg + ".png");
                Picture piclogo = img.CreatePicture();
                
                // Insert a Paragraph into the default Header.
                Paragraph p1 = header_default.InsertParagraph();
                //p1.Append("<html> <body><img src='data:image/png;base64,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' alt='logo' /> </body> </html>");
                p1.AppendPicture(piclogo);
                p1.Alignment = Alignment.right;
                //// Insert a Paragraph into the document.
                //Paragraph p2 = document.InsertParagraph();
                //p2.AppendLine("Hello Document.").Bold();

                

                // Insert a Paragraph into the default Footer.
                Paragraph p3 = footer_default.InsertParagraph();
                p3.Append(footerText).Bold();
                p3.Alignment = Alignment.center;

                // Save all changes to this document.
                document.Save();
                Process.Start("WINWORD.EXE", fileName);
            }// Release this document from memory.



            MessageBox.Show("File saved at D:/" + this.fileNameWithoutExt + ".docx successfully");
           // File.Delete(@"D:\DocLogo.png");
        }

        public class document
        {
            public string sectionID { get; set; }
            public string sectionType { get; set; }
            public string SectionName { get; set; }
            public string sectionContent { get; set; }

        }
    }
}
