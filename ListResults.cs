using System;
using System.Collections.Generic;
using System.Text;
using ArtOfTest.WebAii.Design.Execution;
using System.IO;
using System.Data;
using System.Net.Mail;
using System.Net;
using ArtOfTest.WebAii;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Design.ProjectModel;
using ArtOfTest.WebAii.Design;
using System.Xml;
using System.Xml.Xsl;

namespace TestStudio.Jenkins.Reports
{
    public class ListResults : IExecutionExtension, ITestListExecutionExtension
    {
        string browserType = string.Empty;
        string projectPath = string.Empty;
        string path = string.Empty;
        string environment = string.Empty;

        public void OnBeforeTestListStarted(TestList list)
        {
            browserType = list.Settings.Web.Browser.ToString();
        }

        public void OnAfterTestListCompleted(RunResult result)
        {
            string suiteName = result.Name.ToString();
            string failedCount = result.FailedCount.ToString();
            string passedCount = result.PassedCount.ToString();
            string notRunCount = result.NotRunCount.ToString();
            string totalCount = result.AllCount.ToString();
            string executionTime = (result.EndTime - result.StartTime).Hours.ToString() + " hrs " + (result.EndTime - result.StartTime).Minutes.ToString() + " min " + (result.EndTime - result.StartTime).Seconds.ToString() + " sec";
            string filePath = "C:\\Files\\ResultsXMLFile.xml";
            XmlTextWriter textWriter = new XmlTextWriter(filePath, null);
            textWriter.WriteStartDocument();
            textWriter.WriteComment("Xml file with test suite report");
            textWriter.WriteStartElement("Suite");
            textWriter.WriteAttributeString("title", "UI Test Automation Report");
            textWriter.WriteAttributeString("suitename", suiteName);
            textWriter.WriteAttributeString("total", totalCount);
            textWriter.WriteAttributeString("passed", passedCount);
            textWriter.WriteAttributeString("failed", failedCount);
            textWriter.WriteAttributeString("skipped", notRunCount);
            textWriter.WriteAttributeString("time", executionTime);
            textWriter.WriteAttributeString("browser", browserType);
            List<TestResult> list = result.TestResults;
            foreach (TestResult test in list)
            {
                string description = test.TestDescription;
                string testName = test.TestName;
                string startTime = test.StartTime.ToString();
                string endTime = test.EndTime.ToString();
                string duration = test.Duration.Minutes.ToString() + " min " + test.Duration.Seconds.ToString() + " sec";
                string testResult = test.Result.ToString().ToUpper();
                string testLog = test.Message.ToString();
                textWriter.WriteStartElement("Test");
                textWriter.WriteElementString("Name", testName);
                textWriter.WriteElementString("StartTime", startTime);
                textWriter.WriteElementString("EndTime", endTime);
                textWriter.WriteElementString("Duration", duration);
                textWriter.WriteElementString("Result", testResult);
                textWriter.WriteElementString("LogMessages", testLog);
                textWriter.WriteElementString("TestDescription", description);
                textWriter.WriteEndElement();
            }

            textWriter.WriteEndElement();
            textWriter.WriteEndDocument();
            textWriter.Close();

            string newXmlfile = @"C:\Files\XmlFile.xml";
            string oldXmlFile = filePath;
            CopyXmlDocument(oldXmlFile, newXmlfile);

            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load("C:\\Files\\CSEmailableReport.xslt");
            string htmlFilePath = "C:\\Files\\CSharpTestResultsreport" + (new Random().Next(1000000 - 1) + 1).ToString() + ".html";
            xslt.Transform(newXmlfile, htmlFilePath);

            StreamReader reader = new StreamReader(@"C:\Files\emailProperties.txt");
            Dictionary<string, string> emailProperties = new Dictionary<string, string>();
            while (reader.Peek() >= 0)
            {
                string line = reader.ReadLine();
                emailProperties.Add(line.Split('=')[0], line.Split('=')[1]);
            }
            MailAddress from = new MailAddress(emailProperties["fromAddress"]);

            MailMessage mail = new MailMessage()
            {
                From = from,
                Subject = "Test Results || Pacific || " + suiteName + " || Execution Time : " + executionTime + " || " + result.PassedPercent.ToString() + "% Passed || (" + passedCount + "/" + result.AllCount + ") Passed || Machine : " + result.Machine.NetworkName,
                Body = File.OpenText(htmlFilePath).ReadToEnd()
            };
            string[] toAddresses = emailProperties["toAddresses"].Split(',');
            foreach (string email in toAddresses)
            {
                mail.To.Add(new MailAddress(email));
            }
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "pubsmtp.bedford.progress.com";
            smtp.Port = 25;
            smtp.Credentials = new NetworkCredential("psomasan@progress.com", "Pr@v@ll!k@23");
            smtp.Send(mail);
        }

        public void OnBeforeTestStarted(ExecutionContext executionContext, Test test)
        {
            environment = test.CustomProperty1;
            projectPath = executionContext.DeploymentDirectory;
            Manager.Current.Log.WriteLine("#### Inside Before Test Started Method ####");
        }

        public void OnAfterTestCompleted(ExecutionContext context, TestResult result)
        {
            Manager.Current.Log.WriteLine("#### Inside After Test Completed Method ####");
        }

        public DataTable OnInitializeDataSource(ExecutionContext executionContext)
        {
            return null;
        }

        public void OnStepFailure(ExecutionContext executionContext, AutomationStepResult stepResult)
        {
        }

        private static void CopyXmlDocument(string oldfile, string newfile)
        {
            XmlDocument document = new XmlDocument();
            document.Load(oldfile);
            Console.WriteLine(document.OuterXml);
            if (File.Exists(newfile))
                File.Delete(newfile);
            document.Save(newfile);
        }
    }
}
