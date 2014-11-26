using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using System.Web.Script.Serialization;


namespace DocxCombo
{
    class Question
    {
        public string questionType;
        public int seq;
        public string qAnalysis;
        public string qAnswer;
        public string qBody;
        public string qDocx;
        public string rightOption;

    }

    class PaperCombo
    {
        public string paperName;
        public string paperSubject;
        public int countQuestion;
        public List<Question> questions;

    }


    class Program
    {
        static string path = @"F:\tmp\";


        private static DirectoryInfo download()
        {
            string path = @"F:\tmp\";
            string tempFolderPath = System.Guid.NewGuid().ToString();


            DirectoryInfo dir = System.IO.Directory.CreateDirectory(path + tempFolderPath);

            WebClient webClient = new WebClient();

            string[] urls = {"http://10.60.0.33/download/qd_i2emfd9572.docx",
                             "http://10.60.0.33/download/qd_hygUS37346.docx",
                             "http://10.60.0.33/download/qd_hygUS72295.docx",
                             "http://10.60.0.33/download/qd_hygUSA3382.docx"};

            foreach (string url in urls)
            {
                System.Console.WriteLine(path + tempFolderPath + url.Substring(url.LastIndexOf("/")));

                webClient.DownloadFile(url, path + tempFolderPath + url.Substring(url.LastIndexOf("/")));
            }
            //System.Console.ReadKey();


            return dir;
        }

        private static void compose(DirectoryInfo dir)
        {

            FileInfo[] files = dir.GetFiles();
            Document template = new Document();
            //template.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
            Document tempDoc = new Document();
            foreach (FileInfo file in files)
            {
                tempDoc = new Document(file.FullName);
                tempDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
                template.AppendDocument(tempDoc, ImportFormatMode.KeepSourceFormatting);

            }

            template.Save(System.IO.Path.Combine(dir.FullName, "result.docx"));




            if (false)
            {

                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("hello world!");
                doc.Save(path + "hello.docx");




                Document doc1 = new Document(path + "h1.docx");
                Document doc2 = new Document(path + "h2.docx");
                Document doc3 = new Document();
                doc2.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

                doc1.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);
                //doc1.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                //doc3.AppendDocument(doc2,ImportFormatMode.KeepSourceFormatting);
                doc1.Save(path + "h3.docx");
            }
        }


        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                System.Console.WriteLine(arg);
            }
            if (args.Length > 0)
            {
                string jsonPath = args[0];



                if (!File.Exists(jsonPath))
                {
                    throw new Exception("cannot find json.");
                }
                //读取文件
                using (StreamReader sr = File.OpenText(jsonPath))
                {

                    string jsonStr = sr.ReadToEnd();


                    JavaScriptSerializer jss = new JavaScriptSerializer();
                    PaperCombo pc = jss.Deserialize<PaperCombo>(jsonStr);
                    System.Console.WriteLine(pc.paperName);
                    System.Console.WriteLine(pc.questions.Count);
                    System.Console.WriteLine(pc.questions[0].qBody);
                }
            }


            compose(download());


        }
    }
}
