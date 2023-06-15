import java.time.Duration;
import java.util.List;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotInteractableException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Assignment6Q2 {

	public static void main(String[] args) {
		
		
System.setProperty("webdriver.chrome.driver", "C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		driver.get("https://apps.ualberta.ca/");
		driver.manage().window().maximize();
		String mainw = driver.getWindowHandle();
		Actions act = new Actions(driver);
		List<WebElement> links = driver.findElements(By.xpath("//a[position() > 1]"));
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		
		//Set<String> handle = driver.getWindowHandles();
		
		
		for(WebElement temp:links) {
			
		 act.keyDown(Keys.CONTROL).click(temp).keyUp(Keys.CONTROL).perform();
		 	
		 /*act.keyDown(Keys.CONTROL).perform();
		 temp.click();
		 act.keyUp(Keys.CONTROL).perform();*/
		 
		 Set<String> allw = driver.getWindowHandles();
		
		for(String temp1:allw) {
			if(! mainw.equals(temp1)) {
				
				driver.switchTo().window(temp1);
				
				System.out.println(driver.getCurrentUrl());
				
			   driver.close();		 
		}	   
	}
		 driver.switchTo().window(mainw);

			
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		}
	    
}
