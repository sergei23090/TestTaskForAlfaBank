using System;
using OpenQA.Selenium;
using System.IO;
using ExcelDataReader;
using System.Data;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;

namespace TestTask
{
    class Program
    {
       
        static void Main(string[] args)
        {
            //создаем драйвер для Google Chrome
            IWebDriver driver = new ChromeDriver();
            //Открываем поток для чтения xlsx файла
            FileStream stream = File.Open(@"challenge.xlsx", FileMode.Open, FileAccess.Read);
            //записываем данные из файла в DataTable
            DataTable dt = ExcelToDataTable(stream);
            //Закрываем поток
            stream.Close();
            //Создаем массив строчек, с помощью которого мы будем искать поля для заполнения
            string[] labelNames = {"labelFirstName","labelLastName", "labelCompanyName"
                    , "labelRole", "labelAddress", "labelEmail", "labelPhone" };
            //Переходим по адресу
            driver.Navigate().GoToUrl("https://rpachallenge.com/?lang=EN");
            //Нажимаем на кнопку Start
            driver.FindElement(By.CssSelector("body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button")).Click();
            //Запускаем цикл на ввод значений из xlsx файла в поля на сайте
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count -1; j++)
                {
                    driver.FindElement(By.XPath($"//*[@ng-reflect-name='{labelNames[j]}']")).SendKeys(dt.Rows[i][j].ToString());              
                }
                driver.FindElement(By.XPath("//*[@value='Submit']")).Click();
            }
            //В конце делаем перерыв 3 секунды, чтобы посмотреть успешность прохождения теста
            System.Threading.Thread.Sleep(3000);
            //Закрываем поток и программу
            driver.Close();
            Process.GetCurrentProcess().Kill();
        }

        private static DataTable ExcelToDataTable(FileStream stream)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                return result.Tables[0];
            }
        }

    }
}
