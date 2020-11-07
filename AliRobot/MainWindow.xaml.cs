using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AliRobot
{
    /// <summary>
    /// Interação lógica para MainWindow.xam
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            PesquisaURl();
        }


        IWebDriver driver;

        public void PesquisaURl()
        {
            ChromeDriver driver = new ChromeDriver();
            string Url = "https://pt.aliexpress.com/item/4001119810724.html?spm=a2g0o.productlist.0.0.5ab96c2awRdZIz&s=p&ad_pvid=202009031053563599469359044750002867920_1&algo_pvid=1ab80081-792b-49b5-87b4-31524bacf481&algo_expid=1ab80081-792b-49b5-87b4-31524bacf481-4&btsid=0ab50f4415991556361922742ea5fd&ws_ab_test=searchweb0_0,searchweb201602_,searchweb201603_";
            try
            {
                driver.Navigate().GoToUrl(Url);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//*[@id='root']/div/div[2]/div/div[2]/div[1]/h1")));
                driver.FindElement(By.XPath("//*[@id='root']/div/div[2]/div/div[2]/div[3]/div"));
                audit("PRODUTO INATIVO",Url);

            }
            catch (Exception)
            {
                try
                {
                    var a = driver.FindElement(By.XPath("//*[@id='root']/div/div[2]/div/div[2]/div[4]/div[1]/span")).Text;
                }
                catch (Exception)
                {

                    throw;
                }
                
            }
            

        }
        
        StreamWriter auditwriter = new StreamWriter($@"{AppDomain.CurrentDomain.BaseDirectory}\Logs\audit-{DateTime.Today.ToString("dd-MM-yy")}.txt", true) { AutoFlush = true };
        readonly static object auditObj = new object();


        public void audit(string tag, string texto)
        {
            System.IO.Directory.CreateDirectory($@"{AppDomain.CurrentDomain.BaseDirectory}\Logs");
            lock (auditObj)
            {
                // if (nivelDeAuditoria <= Settings.Default.Auditoria)
                //  {
                auditwriter.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")}\t{tag} >>\t{texto}");
                // }
            }
        }//Escreve no arquivo da auditoria.

        /*
        public void ExcelReader() 
        {
             
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    //var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                }
            }*/

    }
}

    

