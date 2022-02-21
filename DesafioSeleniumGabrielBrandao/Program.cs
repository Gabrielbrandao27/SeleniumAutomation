using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Selenium_Automation
{
	class Program
	{
		//Lê os CEPs da planilha
        public static String GetData(int row, int col)
		{
			String dado;

			excel.Application x1app = new excel.Application();

			excel.Workbook x1Wb = x1app.Workbooks.Open(@"C:\Users\bielf\Desktop\GoLiveTech - Exercício Rafael - Base CEPs.xlsx");

			excel.Worksheet x1Ws = x1Wb.Sheets[1];

			excel.Range x1range = x1Ws.UsedRange;

			dado = Convert.ToString(x1range.Cells[row][col].value2);

			x1app.Quit();
			return dado;
		}

		//Escreve os dados retirados do site dos Correios na planilha Output
		public static void WriteExcel(int row, int col, string value)
        {

			excel.Application x1app = new excel.Application();

			excel.Workbook x1Wb = x1app.Workbooks.Open(@"C:\Users\bielf\Desktop\GoLiveTech - Exercício Rafael - Base CEPs - Output.xlsx");
			//excel.Workbook x1Wb = x1app.Workbooks.Open(@"Seu endereço\Nome do arquivo.xlxs");

			excel.Worksheet x1Ws = x1Wb.Sheets[1];

			excel.Range x1range = x1Ws.UsedRange;

			x1range.Cells[row, col] = value;

			x1Wb.Save();

			Marshal.FinalReleaseComObject(x1Ws);
			x1Ws = null;

			x1app.Quit();
		}

		//Faz as ações de abrir o site, inserir os CEPs e retirar os dados do site para serem escritos na planilha
        static void Main()
		{
			IWebDriver driver = new ChromeDriver(@"C:\Users\bielf\Downloads\chromedriver_win32");
			//IWebDriver driver = new ChromeDriver(@"Seu endereço\chromedriver_win32");

			String colLog; String colBairro; String colUF; String colCEP;
			String logData; String bairroData; String UFData; String cepData;


			for (int i = 2; i <= 5; i++)
            {
				String CEP = GetData(i, 1);

				// Vai abrir o URL
				driver.Url = "https://buscacepinter.correios.com.br/app/endereco/index.php";
				driver.FindElement(By.Id("endereco")).Click();
				driver.FindElement(By.Id("endereco")).Clear();
				driver.FindElement(By.Id("endereco")).SendKeys(CEP);
				driver.FindElement(By.Id("btn_pesquisar")).Click();

				Thread.Sleep(3000);

				//Escreve os campos do corpo da tabela
				IWebElement logradouroID = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/tbody/tr/td[1]"));
				logData = logradouroID.Text;
				WriteExcel(i, 2, logData);        //   //*[@id='resultado-DNEC']/tbody/tr/td[1]

				IWebElement bairroID = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/tbody/tr/td[2]"));
				bairroData = bairroID.Text;
				WriteExcel(i, 3, bairroData);

				IWebElement UFID = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/tbody/tr/td[3]"));
				UFData = UFID.Text;
				WriteExcel(i, 4, UFData);

				IWebElement CEP2ID = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/tbody/tr/td[4]"));
				cepData = CEP2ID.Text;
				WriteExcel(i, 5, cepData);

			}
			//Escreve os campos dos títulos da tabela
			IWebElement logradouro = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/thead/tr/th[1]"));
			colLog = logradouro.Text;	
			WriteExcel(1, 2, colLog);       //  //*[@id="resultado-DNEC"]/thead/tr/th[1]

			IWebElement bairro = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/thead/tr/th[2]"));
			colBairro = bairro.Text;
			WriteExcel(1, 3, colBairro);   

			IWebElement UF = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/thead/tr/th[3]"));
			colUF = UF.Text;
			WriteExcel(1, 4, colUF);

			IWebElement CEP2 = driver.FindElement(By.XPath("//*[@id='resultado-DNEC']/thead/tr/th[4]"));
			colCEP = CEP2.Text;
			WriteExcel(1, 5, colCEP);

			driver.Quit();
		}
	}
}