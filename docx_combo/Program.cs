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
        public string qOptions;
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

        private static DirectoryInfo downloadQuestionDocxs(List<Question> questions)
        {
            string tempFolderPath = System.Guid.NewGuid().ToString();
            DirectoryInfo dir = System.IO.Directory.CreateDirectory(path + tempFolderPath);
            WebClient webClient = new WebClient();
            foreach (Question question in questions)
            {
                webClient.DownloadFile("http://res01.ezxdf.cn/download/" + question.qBody, path + tempFolderPath + "\\" + question.qBody);
                if (question.questionType.Equals("选择题"))
                {
                    webClient.DownloadFile("http://res01.ezxdf.cn/download/" + question.qOptions, path + tempFolderPath + "\\" + question.qOptions);
                }
                webClient.DownloadFile("http://res01.ezxdf.cn/download/" + question.qAnswer, path + tempFolderPath + "\\" + question.qAnswer);
                webClient.DownloadFile("http://res01.ezxdf.cn/download/" + question.qAnalysis, path + tempFolderPath + "\\" + question.qAnalysis);
            }
            return dir;
        }



        private static void composeQuestionDocx(DirectoryInfo dir, List<Question> questions,bool withAnswer)
        {
            Dictionary<string,FileInfo> docxDict = new Dictionary<string, FileInfo>();
            FileInfo[] fileList = dir.GetFiles();
            foreach (FileInfo file in fileList)
            {
                docxDict.Add(file.Name,file);
            }
            Document questionsDocx = new Document();
            Document appenderQuestionDocx = new Document();
            Document appenderOptionDocx = new Document();




            DocumentBuilder docbuilder = null;

            foreach (Question question in questions)
            {

                appenderQuestionDocx = new Document(docxDict[question.qBody].FullName);
                if (question.questionType.Equals("选择题"))
                {
                    appenderOptionDocx = new Document(docxDict[question.qOptions].FullName);
                    appenderOptionDocx.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
                    appenderQuestionDocx.AppendDocument(appenderOptionDocx, ImportFormatMode.KeepSourceFormatting);
                }
                docbuilder = new DocumentBuilder(appenderQuestionDocx);

                docbuilder.MoveToParagraph(0, 0);
                docbuilder.Write(question.seq+1+ ".");
                docbuilder.MoveToDocumentEnd();
                docbuilder.Writeln("");
                docbuilder.Writeln("");
                appenderQuestionDocx.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
                questionsDocx.AppendDocument(appenderQuestionDocx, ImportFormatMode.KeepSourceFormatting);


            }

            if (withAnswer)
            {
                Document answersDocx = new Document();
                Document appenderAnswerDocx = new Document();
                foreach (Question question in questions)
                {
                    appenderAnswerDocx = new Document(docxDict[question.qAnswer].FullName);
                    appenderAnswerDocx.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
                    answersDocx.AppendDocument(appenderAnswerDocx, ImportFormatMode.KeepSourceFormatting);
                }
                questionsDocx.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                questionsDocx.AppendDocument(answersDocx, ImportFormatMode.KeepSourceFormatting);
            }


            questionsDocx.Save(System.IO.Path.Combine(path, "result.docx"));




        }


        static void Main(string[] args)
        {
            bool withAnswerSheet = false;
            foreach (string arg in args)
            {
                System.Console.WriteLine(arg);
            }

            if (args.Length > 0)
            {
                string jsonPath = args[0];
                if (args.Length > 1)
                {
                    withAnswerSheet = Boolean.Parse(args[1]);
                }


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



                composeQuestionDocx(downloadQuestionDocxs(pc.questions), pc.questions, withAnswerSheet);

            }
        }
    }
}
