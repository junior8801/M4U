using System;
using System.IO;
using System.Diagnostics;
using EC = SeleniumExtras.WaitHelpers.ExpectedConditions;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace M4U
{
    class M4U
    {
        static void Main(string[] args)
        {
            string user = "";
            string pass = "";
            string dt1 = DateTime.Now.ToString("dd/MM/yyyy");
            string dt2 = DateTime.Now.ToString("dd/MM/yyyy");
            string rel1 = "https://portaloi.m4u.com.br/ExibeReport.aspx?id=1870";
            //string rel2 = "https://portaloi.m4u.com.br/ExibeReport.aspx?id=1994";
            string downloadFilepath = @"\\server\dir\OI\Relatorios\Parcial_Controle\Base";
            //string downloadFilepath = @"C:\Users\Administrador\Desktop\JUNIOR\Py\M4U";
            
            matarProcessos();
            
            ChromeOptions opt = new ChromeOptions();
            opt.AddUserProfilePreference("download.default_directory", downloadFilepath);
            
            ChromeDriver driver = new ChromeDriver(Directory.GetCurrentDirectory(),opt);               
            driver.Navigate().GoToUrl("https://portaloi.m4u.com.br/Login.aspx");
            Console.WriteLine("Aguardar  a página carregar");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10000));
            Console.WriteLine("Aguardar a página de login carregar");
            wait.Until(EC.ElementIsVisible(By.Id("ctl00_MainContent_LoginButton")));
            driver.FindElementByXPath("/html/body/div[1]/form/div[3]/input[1]").SendKeys (user);             
            driver.FindElementByXPath("/html/body/div[1]/form/div[3]/input[2]").SendKeys (pass);             
            driver.FindElementByXPath("/html/body/div[1]/form/div[3]/input[3]").Click();             
            
            Console.WriteLine("Verificar se existe os arquivos baixados, caso verdadeiro excluir");          
            DirectoryInfo Dir = new DirectoryInfo(downloadFilepath);
            FileInfo[] Files = Dir.GetFiles("AdesoesDetalhado_*.csv", SearchOption.TopDirectoryOnly);
            foreach (FileInfo Arq in Files){                
                File.Delete(Arq.FullName);
            }
            Files = Dir.GetFiles("Adesoes por periodo*.csv", SearchOption.TopDirectoryOnly);
            foreach (FileInfo Arq in Files){                
                File.Delete(Arq.FullName);
            }

            driver.Navigate().GoToUrl(rel1);
            Console.WriteLine("Aguardar a página do controle cartão carregar");
            wait.Until(EC.ElementIsVisible(By.Id("ctl00_MainContent_RPV01_ctl04_ctl00")));
            Console.WriteLine("Preencher a data");
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/div/input[1]").SendKeys (dt1);     
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[5]/div/div/input[1]").SendKeys (dt2);     
            Console.WriteLine("Gerar o relatório");
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[3]/table/tbody/tr/td/input").Click();                                 
            
            Console.WriteLine("Esperar imagem de carregamento sumir ou o relatório carregar");          
            wait.Until(EC.InvisibilityOfElementLocated(By.Id("ctl00_MainContent_RPV01_AsyncWait_Wait")));
            Console.WriteLine("Baixar base controle Cartão");          
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[4]/td/div/div/div[2]/table/tbody/tr/td/div[1]/table/tbody/tr/td/a/img[2]").Click();                                                     
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[4]/td/div/div/div[2]/table/tbody/tr/td/div[2]/div[3]/a").Click();                                                                         
            
            //boleto arquivo
            Console.WriteLine("Ir para página do arquivo controle boleto");          
            driver.Navigate().GoToUrl("https://portaloi.m4u.com.br/ListaArquivosS3.aspx?key=Oi/OI_ControleDigital/OUT/AdesoesDetalhadoHoraHora/");
            Console.WriteLine("Baixar base controle boleto");          
            driver.FindElementByXPath("/html/body/form/div[3]/div/div[2]/div/div[1]/div[2]/ul/li[1]/a").Click();                                                     
            
            Console.WriteLine("Excluir as últimas bases se existirem");          
            if(File.Exists(downloadFilepath+@"\CONTROLE_CARTAO.CSV")){
                File.Delete(downloadFilepath+@"\CONTROLE_CARTAO.CSV");            
            }              
            if(File.Exists(downloadFilepath+@"\CONTROLE_BOLETO.CSV")){
                File.Delete(downloadFilepath+@"\CONTROLE_BOLETO.CSV");            
            }

            Console.WriteLine("Verificar se terminou o bownload da base controle boleto");          
            Dir = new DirectoryInfo(downloadFilepath);
            Files = Dir.GetFiles("AdesoesDetalhado_*.csv", SearchOption.TopDirectoryOnly);

            while (Files.Length <= 0)
            {
                Console.WriteLine("Ainda não terminou, esperar 10 segundos"); 
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(10));
                Files = Dir.GetFiles("AdesoesDetalhado_*.csv", SearchOption.TopDirectoryOnly);
            }
            Console.WriteLine("Renomear arquivos"); 
            foreach (FileInfo Arq in Files){
                    string FilePath = Arq.FullName;
                    string FileName = Arq.Name;
                    Console.WriteLine(FileName);
                    File.Move(downloadFilepath+@"\"+FileName, downloadFilepath+@"\CONTROLE_BOLETO.CSV");                    
            }
            File.Move(downloadFilepath+@"\Adesoes por periodo.CSV", downloadFilepath+@"\CONTROLE_CARTAO.CSV");

            Console.WriteLine("Finalizar execução");          
            driver.Quit();            
                    
        }
        public static void matarProcessos(){            
            foreach (Process process in Process.GetProcessesByName("chromedriver")){
                process.Kill();
            }
            foreach (Process process in Process.GetProcessesByName("chrome")){
               process.Kill();
            }
        }
    }
}
