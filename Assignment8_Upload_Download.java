import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Assignment8_Upload_Download {

	public static void main(String[] args) throws AWTException {
		
	
System.setProperty("webdriver.chrome.driver", "C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
	
		
		
		driver.get("https://demoqa.com/upload-download");
		driver.manage().window().maximize();
		
		Robot r = new Robot();
		driver.findElement(By.xpath("//a[@id=\"downloadButton\"]")).click();
		r.delay(5000);
		
		//driver.findElement(By.xpath("//input[@id=\"uploadFile\"]")).click();
		
		r.keyPress(KeyEvent.VK_TAB);
		r.keyRelease(KeyEvent.VK_TAB);
		r.keyPress(KeyEvent.VK_ENTER);
		r.keyRelease(KeyEvent.VK_ENTER);
		r.delay(5000);
		r.keyPress(KeyEvent.VK_S);
		r.keyRelease(KeyEvent.VK_S);
		r.keyPress(KeyEvent.VK_C);
		r.keyRelease(KeyEvent.VK_C);
		r.delay(5000);
		r.keyPress(KeyEvent.VK_DOWN);
		r.keyRelease(KeyEvent.VK_DOWN);
		r.keyPress(KeyEvent.VK_ENTER);
		r.keyRelease(KeyEvent.VK_ENTER);
		
		

	}

}
