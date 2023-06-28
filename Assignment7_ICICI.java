import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

public class Assignment7_ICICI {
	static WebDriver driver;

	public static void main(String[] args) throws InterruptedException {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");

		driver = new ChromeDriver();
		driver.get("https://www.icicibank.com/#");
		driver.manage().window().maximize();
		element("//a[@id=\"top-header-item-2\"]");
		element("//a[@class=\"agree\"]");
		Thread.sleep(5000);
		Actions act = new Actions(driver);
		act.moveToElement(driver.findElement(By.xpath("//span[@class=\"n-icon\"]/img[@src='/content/dam/icicibank/icici-assets/nri-banking/homepage/savings-ac.png']"))).click().build().perform();
		act.moveToElement(driver.findElement(By.xpath("//div[@class='card-item billcard-item slick-slide slick-current slick-active']//li[1]//div[2]//a[1]"))).click().build().perform();
		Thread.sleep(5000);
		
		ChromeOptions options = new ChromeOptions();
        options.addArguments("--enable-features=PushMessaging");
		
       

	}
		
		public static void element(String xpath) {

			driver.findElement(By.xpath(xpath)).click();

		

	}

}
