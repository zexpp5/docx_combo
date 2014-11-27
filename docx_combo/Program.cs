using System;
using System.Collections;
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
    class Question:IComparable
    {
        public string questionType;
        public int seq;
        public string qAnalysis;
        public string qAnswer;
        public string qBody;
        public string qDocx;
        public string rightOption;

        public int CompareTo(object obj)
        {
            int res = 0;
            try
            {
                Question target = (Question)obj;
                if (this.seq > target.seq)
                {
                    res = 1;
                }
                if (this.seq < target.seq)
                {
                    res = -1;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("comparation exception.",ex.InnerException);
            }

            return res;
        }
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
        static string path = @"D:\tmp\";


        private static FileInfo[] download(string[] urls)
        {
            string tempFolderPath = System.Guid.NewGuid().ToString();


            DirectoryInfo dir = System.IO.Directory.CreateDirectory(path + tempFolderPath);

            WebClient webClient = new WebClient();


            foreach (string url in urls)
            {
                System.Console.WriteLine(url);

                System.Console.WriteLine(path + tempFolderPath +"\\"+ url);

                webClient.DownloadFile("http://res01.ezxdf.cn/download/"+url, path + tempFolderPath +"\\"+ url);
            }
            //System.Console.ReadKey();

            System.Console.WriteLine(dir.ToString());
            return dir.GetFiles();
        }

        private static void compose(FileInfo[] files)
        {

            Document template = new Document();
            //template.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
            Document tempDoc = new Document();
            foreach (FileInfo file in files)
            {
                tempDoc = new Document(file.FullName);
                tempDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
                template.AppendDocument(tempDoc, ImportFormatMode.KeepSourceFormatting);

            }

            template.Save(System.IO.Path.Combine(path, "result.docx"));




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

                PaperCombo pc = null;
                using (StreamReader sr = File.OpenText(jsonPath))
                {

                    string jsonStr = sr.ReadToEnd();


                    JavaScriptSerializer jss = new JavaScriptSerializer();
                    pc = jss.Deserialize<PaperCombo>(jsonStr);
                    System.Console.WriteLine(pc.paperName);
                    System.Console.WriteLine(pc.questions.Count);
                    System.Console.WriteLine(pc.questions[0].qBody);
                }

                

                pc.questions.Sort();

                int paperSize = pc.questions.Count;

                string[] qBodyPaths = new string[paperSize];
                string[] qAnswerPaths = new string[paperSize];

                for (int i = 0; i < paperSize; i++)
                {
                    qBodyPaths[i] = pc.questions[i].qBody;
                    qAnswerPaths[i] = pc.questions[i].qAnswer;
                }




                compose(download(qBodyPaths));

            }
        }
    }
}
