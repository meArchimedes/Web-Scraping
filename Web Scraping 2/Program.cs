// See https://aka.ms/new-console-template for more information

using HtmlAgilityPack;
using CsvHelper;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Support.UI;
using IronXL;
class Program
{
    public static string? currentShippingCountry;
    static async Task Main(string[] args)
    {
        var html = await GetHtml("https://www.horeca.com/en/categorie/3621/combi-steamers");
        List<string> pageLinks = GetHtmlForPageLinks(html);
        List<string> prices = ParseHtmlForPrices(html);
        List<string> availability = ParseHtmlForItemAvailability(html);
        List<string> itemIds = ParseHtmlForItemIds(html);
        List<string> links = ParseIdsToLinks(itemIds);

        var list = GetDynamicDataFromLinks(links);
        fillExcelSheet(new List<List<string>> { links, itemIds, list.First(), prices, availability, list.Last() });

    }
        private static Task<string> GetHtml(string link)
    {
        var client = new HttpClient();
        return client.GetStringAsync(link);
    }
    private static List<string> GetHtmlForPageLinks(string html)
    {
        List<string> PageLinks = new();
        HtmlDocument htmlDoc = new();
        htmlDoc.LoadHtml(html);
        var pageLinks =
            htmlDoc
            .DocumentNode
            .SelectNodes("//a[(contains(@class, 'page-link'))]");
        foreach (var link in pageLinks)
        {
            PageLinks.Add(link.InnerHtml);
        }
        return PageLinks;
    }
    private static List<string> ParseHtmlForPrices(string html)
    {
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);
        List<string> prices = new();
        var SteamerPrices =
                  htmlDoc
                  .DocumentNode
                  .SelectNodes("//span[(contains(@class, 'product-price'))]");
        foreach (var item in SteamerPrices)
        {
            prices.Add(item.InnerHtml.Substring(0, item.InnerHtml.IndexOf(",")));
        }
        return prices;
    }
    private static List<string> ParseHtmlForItemIds(string html)
    {
        string id = string.Empty;
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);
        List<string> itemId = new();
        var ItemIds =
            htmlDoc
            .DocumentNode
            .SelectNodes("//input[(contains(@id, 'product-comparison'))]");
        foreach (var item in ItemIds)
        {

            var matches = Regex.Matches(item.Id, @"\d+");
            foreach (Match match in matches)
            {
                id += match;
            }
            itemId.Add(id);
            id = string.Empty;
        }
        return itemId;
    }
    private static List<string> ParseIdsToLinks(List<string> itemIds)
    {
        List<string> links = new();
        foreach (string id in itemIds)
        {
            links.Add("https://www.horeca.com/nl/product/" + id);
        }
        return links;
    }
    private static List<string> ParseHtmlForItemAvailability(string html)
    {
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);
        List<string> availability = new();
        var Availability =
                    htmlDoc
                    .DocumentNode
                    .SelectNodes("//div[(contains(@class, 'mb-2 fw-bold'))]").Descendants();
        int i = 1;
                foreach (var item in Availability)
                {
                    if (!item.HasChildNodes && !item.Equals("") && i%3==0)
                    {
                    availability.Add(item.InnerHtml.Trim());
                    }
            i++;
                }
        return availability;
    }
    private static List<List<string>> GetDynamicDataFromLinks(List<string> links)
    {
        List<string> deliveryTime = new();
        List<string> shippingCosts = new();
        WebDriver driver = new ChromeDriver();
        bool hasEntered = false;
        bool flag = false;
        foreach (var item in links)
        {
            //Go to url of item
            try
            {
                driver.Navigate().GoToUrl(item);

                if (!hasEntered)
                {
                    driver.Manage().Window.Maximize();
                    if (!flag)
                    {
                        flag = CloseLanguageTab(flag, driver);
                    }
                    hasEntered = ChangeShippingCountry(driver, hasEntered);
                }
            } 
            catch
            {
                hasEntered = false;
            }
            shippingCosts = GetShippingFeeData(driver, shippingCosts);
            deliveryTime = GetDeliveryData(driver, deliveryTime);
        }
        return new List<List<string>> { deliveryTime, shippingCosts };
    }
    private static bool CloseLanguageTab(bool flag, WebDriver driver)
{
    try
    {
        driver.FindElement(By.XPath("/html/body/div[6]/div/div/div[1]/button")).Click();
        Task.Delay(3000).Wait();
        flag = true;
    }
    catch
    {

    }
    return flag;
}
    public static bool ChangeShippingCountry(WebDriver driver, bool hasEntered)
    {
    currentShippingCountry = driver.FindElement(By.XPath("//strong[(contains(@class, 'shipping-costs-country'))]")).Text.ToString();
    Task.Delay(4000).Wait();
        if (currentShippingCountry.ToString() != null && !currentShippingCountry.ToString().ToLower().Contains("ne") && !currentShippingCountry.ToString().ToLower().Contains("land"))
        {
            //Change shipping country to 'Netherlands' if not set

            //Click on Shipping country dropdown
            driver.FindElement(By.XPath("/html/body/div[3]/div[10]/form/div[1]/div[2]/div/div/div[1]/div[3]/div[2]/div[1]/i")).Click();
            Task.Delay(3000).Wait();

            //Click on country button
            driver.FindElement(By.Id("shippingCountryId")).Click();
            Task.Delay(2000).Wait();

            //select "Netherlands"
            driver.FindElement(By.XPath("/html/body/div[3]/div[10]/form/div[1]/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[1]/select/option[19]")).Click();
            Task.Delay(2000).Wait();

            //Enter zip code
            driver.FindElement(By.XPath("/html/body/div[3]/div[10]/form/div[1]/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[2]/input")).SendKeys("1234");

            //Click "calculate"
            driver.FindElement(By.XPath("/html/body/div[3]/div[10]/form/div[1]/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/button")).Click();
            Task.Delay(5000).Wait();
            hasEntered = true;
        }
            return hasEntered;
    }
    private static List<string> GetShippingFeeData(WebDriver driver, List<string> shippingCosts)
    {
        try
        {
            shippingCosts.Add(driver.FindElement(By.XPath("//strong[(contains(@class, 'shipping-costs-price'))]")).Text);
        }
        catch
        {
            shippingCosts.Add("Not_Found");
        }
        return shippingCosts;
    }
    private static List<string> GetDeliveryData(WebDriver driver, List<string> deliveryTime)
    {
        try
        {
            deliveryTime.Add(driver.FindElement(By.XPath("//div[(contains(@id, 'deliveryTime'))]")).Text);
        }
        catch
        {
            deliveryTime.Add("Not_Found");
        }
        return deliveryTime;
    }

    private static void fillExcelSheet(List<List<string>> list)
    {
        WorkBook oWB = WorkBook.Create();
        var oSheet = oWB.CreateWorkSheet("combi_steamers_data");

        oSheet["A1"].Value = "Product Web Address";
        oSheet["B1"].Value = "Product Number";
        oSheet["C1"].Value = "Delivery Time";
        oSheet["D1"].Value = "Price";
        oSheet["E1"].Value = "Availability";
        oSheet["F1"].Value = "Shipping Costs";

        List<string> cells = new() { }; 
        char j = 'A';
        foreach (var item in list)
        {
            for(int i = 2; i <= item.Count+1; i++)
                    {
                        oSheet[j+i.ToString()].Value = item[i-2];
                    }
            j = (Char)(Convert.ToUInt16(j) + 1); ;
        }
        
        oWB.SaveAs(@"C:\\Users\\User\\Downloads\\WebScraping-Excel.xls");
        oWB.Close();
    }
    }