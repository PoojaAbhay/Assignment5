import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class BankOfAmerica {
	static WebDriver driver;

	public static void main(String[] args) throws InterruptedException {
		
		
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");

		driver = new ChromeDriver();
		driver.get("https://www.bankofamerica.com/");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		
		element("//a[@id=\"navCreditCards\"]").click();
		Thread.sleep(5000);
		element("//a[@id=\"cashRewardsCards\"]").click();
		element("//a[@id=\"home_4063779~A4~en_US\"]").click();
		
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("document.body.style.zoom='25%'");
		
		Thread.sleep(5000);
		
		element("//input[@id=\"customerFirstName\"]").sendKeys("Bank");
		element("//input[@id=\"customerMiddleName\"]").sendKeys("Of");
		element("//input[@id=\"customerLastName\"]").sendKeys("America");
		
		Select sel = new Select(element("//select[@id=\"customerNameSuffix\"]"));
		sel.selectByVisibleText("Sr");
		
		element("//input[@id=\"customerResidentialAddressOne\"]").sendKeys("701 Tyler Ave Warrensburg");
		
		element("//input[@id=\"customerAddressCity\"]").sendKeys("Missouri");
		
		Select sele = new Select(element("//select[@id=\"customerAddressState\"]"));
		sele.selectByVisibleText("Missouri");
		
		element("//input[@id=\"customerAddressInput\"]").sendKeys("63151");
		element("//input[@id=\"customerPrimaryPhoneNumber\"]").sendKeys("3469757585");
		Thread.sleep(5000);
		element("//input[@id=\"phoneNumberMobile\"][@value=\"CellularPhone\"]").click();
		element("//input[@id=\"customerEmailAddress\"]").sendKeys("abc@gmail.com");
		
		
		
	}

	public static WebElement element(String xpath) {

		return driver.findElement(By.xpath(xpath));

	}

}
