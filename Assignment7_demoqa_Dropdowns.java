import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Assignment7_demoqa_Dropdowns {

	static WebDriver driver;

	public static void main(String[] args) throws InterruptedException {

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");

		driver = new ChromeDriver();
		driver.get("https://demoqa.com/select-menu");
		driver.manage().window().maximize();

		Thread.sleep(5000);
		// Select Value
		element("//div[@id = \"withOptGroup\"]");

		Actions act = new Actions(driver);
		act.sendKeys(Keys.DOWN).sendKeys(Keys.DOWN).sendKeys(Keys.DOWN).sendKeys(Keys.ENTER).build().perform();

		// Select one

		element("//div[@id=\"selectOne\"]");
		act.sendKeys("P").sendKeys(Keys.ENTER).build().perform();

		// Old style select menu

		Select s = new Select(driver.findElement(By.xpath("//select[@id=\"oldSelectMenu\"]")));
		s.selectByIndex(4);

		// Multiselect drop down
		element("//*[@id=\"selectMenuContainer\"]/div[7]/div/div/div/div[1]");
		act.sendKeys("G").sendKeys(Keys.ENTER).sendKeys("R").sendKeys(Keys.ENTER).sendKeys("B").sendKeys(Keys.ENTER)
				.build().perform();
		
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("document.body.style.zoom='50%'");

		// Standard multi select
		Select sel = new Select(driver.findElement(By.xpath("//select[@id=\"cars\"]")));
		sel.selectByVisibleText("Volvo");
		sel.selectByVisibleText("Saab");
		sel.selectByVisibleText("Opel");
		sel.deselectAll();
		

		
	}

	public static void element(String xpath) {

		driver.findElement(By.xpath(xpath)).click();

	}

}
