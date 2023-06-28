import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Frames {

	public static void main(String[] args) throws InterruptedException {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		driver.get("https://demoqa.com/frames");
		driver.manage().window().maximize();
		
		// Freames
		
		driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id=\"frame1\"]")));
		System.out.println(driver.findElement(By.xpath("//h1[@id='sampleHeading']")).getText());
		driver.switchTo().defaultContent();
		driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id=\"frame2\"]")));
		System.out.println(driver.findElement(By.xpath("//h1[@id='sampleHeading']")).getText());
		
	
		// Nested frames
		
		driver.get("https://demoqa.com/nestedframes");
		driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id=\"frame1\"]")));
		System.out.println(driver.findElement(By.xpath("//body[text()='Parent frame']")).getText());
		driver.switchTo().frame(driver.findElement(By.xpath("//body[text()='Parent frame']/child::iframe")));
		System.out.println(driver.findElement(By.xpath("//p[normalize-space()='Child Iframe']")).getText());
		
	
	}

}
