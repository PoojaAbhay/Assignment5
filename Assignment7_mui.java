import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Assignment7_mui {

	
		
		
		static WebDriver driver;

		public static void main(String[] args) throws InterruptedException {
			System.setProperty("webdriver.chrome.driver",
					"C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");

			driver = new ChromeDriver();
			driver.get("https://mui.com/material-ui/react-autocomplete/");
			driver.manage().window().maximize();
			driver.findElement(By.xpath("//input[@id='combo-box-demo']")).click();
			driver.findElement(By.xpath("//input[@id='combo-box-demo']")).sendKeys("2001");
			Actions act = new Actions(driver);
			act.sendKeys(Keys.DOWN).sendKeys(Keys.ENTER).build().perform();
			
			

	}

}
