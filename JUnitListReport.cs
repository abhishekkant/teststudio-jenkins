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
    public class CustomReporter : IExecutionExtension, ITestListExecutionExtension
    {
        string browserType = string.Empty;
        string projectPath = string.Empty;
        string path = string.Empty;
        string environment = string.Empty;
        string resultsPath = string.Empty;
        string listName = string.Empty;

        public void OnBeforeTestListStarted(TestList list)
        {
            
            browserType = list.Settings.Web.Browser.ToString();
            listName = list.TestListName;
            string date = DateTime.Now.Date.ToShortDateString().Replace("/", "");
           // resultsPath = list.ProjectPath;
           // resultsPath = "C:\\TelerikReports\\" + date;
            //System.IO.Directory.CreateDirectory(resultsPath);
        }

        public void OnAfterTestListCompleted(RunResult result)
        {

            try {
                string path = resultsPath + "\\ConsolidatedReport.xml";
                string passedCount = result.PassedCount.ToString();
                string failedCount = result.FailedCount.ToString();
                string notRunCount = result.NotRunCount.ToString();
                string totalCount = result.AllCount.ToString();
                string runTime = (result.EndTime - result.StartTime).ToString();

                XmlTextWriter textWriter = new XmlTextWriter(path, null);
                textWriter.WriteStartDocument();
                textWriter.WriteStartElement("testsuites");
                textWriter.WriteStartElement("testsuite");
                textWriter.WriteAttributeString("name", listName);
                textWriter.WriteAttributeString("tests", totalCount.ToString());
                textWriter.WriteAttributeString("failures", failedCount.ToString());
                textWriter.WriteAttributeString("skipped", notRunCount.ToString());
                textWriter.WriteAttributeString("time", runTime);

                List<TestResult> list = result.TestResults;

                foreach (TestResult test in list)
                {
                    string testName = test.TestName;
                    string duration = test.Duration.Minutes.ToString() + " min " + test.Duration.Seconds.ToString() + " sec";
                    string testResult = test.Result.ToString().ToUpper();
                    string testLog = test.Message.ToString();

                    textWriter.WriteStartElement("testcase");
                    textWriter.WriteAttributeString("classname", testName);
                    textWriter.WriteAttributeString("name", testName);
                    textWriter.WriteAttributeString("time", test.Duration.ToString());
                    if (test.Result.ToString().Equals("Fail"))
                    {
                        textWriter.WriteStartElement("failure");
                        textWriter.WriteAttributeString("message", "FAILED");
                        textWriter.WriteEndElement();
                        textWriter.WriteElementString("system-out", testLog);
                    }
                    textWriter.WriteEndElement();
                }

                textWriter.WriteEndElement();
                textWriter.WriteEndElement();
                textWriter.WriteEndDocument();
                textWriter.Close();
            }
            catch (Exception ex)
            {
                Manager.Current.Log.WriteLine(ex.Message  + ex.Source + ex.StackTrace);
            }
        }

        public void OnBeforeTestStarted(ExecutionContext executionContext, Test test)
        {
            string envJob = Environment.GetEnvironmentVariable("JOB_NAME");
            string envBldNum =  Environment.GetEnvironmentVariable("BUILD_NUMBER");
            environment = test.CustomProperty1;
            projectPath = executionContext.DeploymentDirectory;
            resultsPath = executionContext.DeploymentDirectory + @"\" + envJob + envBldNum; 
            Manager.Current.Log.WriteLine("#### Inside Before Test Started Method ####");
            Manager.Current.Log.WriteLine(projectPath);
            Manager.Current.Log.WriteLine(envJob + envBldNum);
        }

        public void OnAfterTestCompleted(ExecutionContext context, TestResult result)
        {
            //Manager.Current.Log.WriteLine("#### Start Inside After Test Completed Method ####");
            //string directoryName = resultsPath + "\\" + result.TestName;
            //System.IO.Directory.CreateDirectory(directoryName);
            //var filePath = directoryName + "\\" + result.TestName + ".xml";
            //int totalTests = 1;
            //int failures = 0;
            //int skipped = 0;
            //string logMessage = null;
            //string runTime = (result.EndTime - result.StartTime).ToString();
            //if (result.Result.ToString().Equals("Fail"))
            //{
            //    logMessage = result.Message.ToString();
            //    failures += 1;
            //}

            //XmlTextWriter textWriter = new XmlTextWriter(filePath, null);
            //textWriter.WriteStartDocument();
            //textWriter.WriteComment("Xml file with test suite report");
            //textWriter.WriteStartElement("testsuite");
            //textWriter.WriteAttributeString("name", listName);
            //textWriter.WriteAttributeString("tests", totalTests.ToString());
            //textWriter.WriteAttributeString("failures", failures.ToString());
            //textWriter.WriteAttributeString("skipped", skipped.ToString());
            //textWriter.WriteAttributeString("time", runTime);
            //textWriter.WriteStartElement("testcase");
            //textWriter.WriteAttributeString("classname", result.TestName);
            //textWriter.WriteAttributeString("name", result.TestName);
            //textWriter.WriteAttributeString("time", runTime);
            //if (result.Result.ToString().Equals("Fail"))
            //{
            //    textWriter.WriteStartElement("failure");
            //    textWriter.WriteAttributeString("message", "FAILED");
            //    textWriter.WriteEndElement();
            //    textWriter.WriteElementString("system-out", logMessage);
            //}
            //textWriter.WriteEndElement();
            //textWriter.WriteEndElement();
            //textWriter.WriteEndDocument();
            //textWriter.Close();
            //totalTests++;
            //Manager.Current.Log.WriteLine("#### End Inside After Test Completed Method ####");
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
