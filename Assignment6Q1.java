import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.opentelemetry.exporter.logging.SystemOutLogRecordExporter;

public class Assignment6Q1 {

	public static void main(String[] args) {
		
System.setProperty("webdriver.chrome.driver", "C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		driver.get("https://demoqa.com/browser-windows");
		driver.manage().window().maximize();
		String main = driver.getWindowHandle();
		System.out.println(main);
		driver.findElement(By.xpath("//button[@id='tabButton']")).click();
		driver.findElement(By.xpath("//button[@id='windowButton']")).click();
		driver.findElement(By.xpath("//button[@id='messageWindowButton']")).click();
		
		Set<String> allw = driver.getWindowHandles();
		for(String Temp:allw) {
			
			System.out.println(Temp);
			
			driver.switchTo().window(Temp);
			
			if(! main.equals(Temp)) {
				
				
				driver.close();
			}
			
			driver.switchTo().window(main);
		}
		
		driver.quit();
	}

}
