using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;

namespace seleniumTest
{
    class MainProcess
    {
        public Form MyForm { get; set; }
        public RichTextBox MyRichTextBox { get; set; }
        public String FileName { get; set; }
        public String Key { get; set; }
        public bool RankEnabled { get; set; }
        public String Rank { get; set; }
        public bool NPEnabled { get; set; }
        public String NP { get; set; }
        public bool MyBuys { get; set; }

        private const String ACCESS_URL = "chrome-extension://oogdoioldgknmlmaaekjfeengjhiekde/popup.html";
        private const String OPTIONS_URL = "chrome-extension://oogdoioldgknmlmaaekjfeengjhiekde/options.html";

        private IWebDriver driver;

        List<int> before;
        List<int> after;

        public void Run()
        {
            Log("Starting scraping job on thread ID " + Thread.CurrentThread.ManagedThreadId);

            var options = new ChromeOptions();
            var driverService = ChromeDriverService.CreateDefaultService();

            before = Process.GetProcessesByName("chrome").Select(p => p.Id).ToList();

            options.AddExtensions("OA-XRAY.crx");
            options.AddArguments("-disable-extensions-http-throttling");
            driver = new ChromeDriver(driverService ,options, TimeSpan.FromMinutes(5));

            Log("Trying to access extension on thread ID: " + Thread.CurrentThread.ManagedThreadId);

            FillData();
        }

        private void FillData()
        {
            var fileContent = File.ReadAllLines(FileName).ToList();
            var path = new FileInfo(FileName).Directory.FullName;
            int actualPID = 0;

            File.WriteAllText(Path.Combine(path, "completed.txt"), String.Empty);

            Log("Links loaded");

            after = Process.GetProcessesByName("chrome").Select(p => p.Id).ToList();

            Log("Setting extension options");

            driver.Url = OPTIONS_URL;
            driver.Navigate();

            Thread.Sleep(2000);

            driver.FindElement(By.Id("keyname")).Clear();
            driver.FindElement(By.Id("keyname")).SendKeys(Key);
            driver.FindElement(By.Id("save")).Click();

            Log("Options set!");
            
            int groupsCounter = (fileContent.Count % 10 == 0) ? fileContent.Count / 10 : (fileContent.Count / 10) + 1;
            int counter = 1;
            Log(String.Format("{0} link groups will be processed", groupsCounter));

            while (fileContent.Count != 0)
            {
                try
                {
                    if (actualPID == 0)
                    {
                        Log("Attempting to get Chrome scraping window");
                        var seleniumPIDs = after.Except(before).ToList();

                        foreach (var process in seleniumPIDs)
                        {
                            Log($"Trying to open extension on PID {process}");

                            try
                            {
                                TestStack.White.Application app = TestStack.White.Application.Attach(process);

                                List<Window> windows = app.GetWindows();

                                Window window = windows[0];

                                TestStack.White.UIItems.Button button = window.Get<TestStack.White.UIItems.Button>("OA XRAY");
                                button.Click();

                                actualPID = process;

                                break;
                            }
                            catch (Exception)
                            {
                                Log($"Skipping PID {process}, trying on next PID");
                            }
                        }

                        Log($"Selenium window running on PID {actualPID}");
                    }
                    else
                    {
                        TestStack.White.Application app = TestStack.White.Application.Attach(actualPID);

                        List<Window> windows = app.GetWindows();

                        Window window = windows[0];

                        TestStack.White.UIItems.Button button = window.Get<TestStack.White.UIItems.Button>("OA XRAY");
                        button.Click();
                    }

                    Log(String.Format("Processing group {0} / {1}", counter++, groupsCounter));

                    driver.SwitchTo().Window(driver.WindowHandles[0]).Close();

                    driver.SwitchTo().Window(driver.WindowHandles[0]);

                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5000));

                    var compare = driver.FindElement(By.Id("iscomparestore"));

                    if (!compare.Selected)
                        compare.Click();

                    var links = driver.FindElement(By.Id("linkstablabel"));
                    links.Click();

                    List<string> temp = new List<string>();
                    var textArea = driver.FindElement(By.TagName("textarea"));

                    textArea.Clear();

                    for (int i = 0; i < 10; i++)
                    {
                        try
                        {
                            temp.Add(fileContent[0]);
                            fileContent.RemoveAt(0);
                        }
                        catch (Exception)
                        {
                            break;
                        }
                    }

                    foreach (var url in temp)
                    {
                        textArea.SendKeys(url);
                        textArea.SendKeys("\n");
                    }

                    driver.FindElement(By.Id("linksbutton")).Click();

                    Log("Waiting for the page to finish processing records");
                    while (driver.FindElement(By.XPath("//span[@class='progress']")).GetAttribute("style") != "width: 100%;")
                    {
                        Thread.Sleep(1000);
                    }

                    Log("Downloading records");

                    js.ExecuteScript("toCSV()");

                    using (var sw = new StreamWriter(Path.Combine(path, "completed.txt"), true))
                    {
                        foreach (var link in temp)
                        {

                            sw.WriteLine(link);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogException(ex);
                }
            }

            driver.Quit();
            Log("Process finished on thread ID " + Thread.CurrentThread.ManagedThreadId);
        }

        private void LogException(Exception ex)
        {
            MyForm.Invoke(new Action(
                delegate()
                {
                    MyRichTextBox.Text += ex.Message + Environment.NewLine;
                    MyRichTextBox.Text += ex.StackTrace + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate()
                {
                    MyForm.Refresh();
                }));
        }

        private void Log(String text)
        {
            MyForm.Invoke(new Action(
                delegate()
                {
                    MyRichTextBox.Text += text + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate()
                {
                    MyForm.Refresh();
                }));
        }
    }
}
