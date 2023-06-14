import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.Color;

public class Q2 {

	public static void main(String[] args) throws InterruptedException {
		
		
System.setProperty("webdriver.chrome.driver", "C:\\Users\\visit\\eclipse-workspace\\SeleniumPractice\\Driver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
	
		
		//1)  https://demoqa.com/dragabble all tabs.
		//1.1 Tab1 - Simple
		
		driver.get("https://demoqa.com/dragabble");
		driver.manage().window().maximize();
		Actions act = new Actions(driver);
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@id=\"dragBox\"]")), 100, 100).perform();
		
		//1.2 Tab2 - Axis Restricted
		
		driver.findElement(By.xpath("//a[@id=\"draggableExample-tab-axisRestriction\"]")).click();
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@id=\"restrictedX\"]")), 100, 0).perform();
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@id=\"restrictedY\"]")), 0, 100).perform();
		
		//1.3 Tab3- Container Restricted
		
		
		driver.findElement(By.xpath("//a[@id=\"draggableExample-tab-containerRestriction\"]")).click();
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@class=\"draggable ui-widget-content ui-draggable ui-draggable-handle\"]")), 50, 30).perform();
		act.dragAndDropBy(driver.findElement(By.xpath("//span[@class=\"ui-widget-header ui-draggable ui-draggable-handle\"]")), 0, 50).perform();
		
		//1.4 Tab4 - Cursor style
		
		driver.findElement(By.xpath("//a[@id=\"draggableExample-tab-cursorStyle\"]")).click();
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@id=\"cursorTopLeft\"]")), 100, 50).perform();	
		act.dragAndDropBy(driver.findElement(By.xpath("//div[@id=\"cursorCenter\"]")), 100, 30).perform();	
		act.dragAndDropBy(driver.findElement(By.xpath("//span[@class=\"bottom-text\"]")), 100, 50).perform();
			
		
		
		//2) https://demoqa.com/droppable all tabs
		
		driver.get("https://demoqa.com/droppable");
		driver.manage().window().maximize();
		
		//Tab1- Simple
		
		act.clickAndHold(driver.findElement(By.xpath("//div[@id=\"draggable\"]"))).moveToElement(driver.findElement(By.xpath("//div[@id=\"droppable\"]"))).perform();
		act.moveByOffset(10, 10).perform();
		String s = driver.findElement(By.xpath("//div[@id=\"droppable\"]")).getCssValue("background-color");
		System.out.println("Background colour of the bloack changes to - " + Color.fromString(s).asHex());
		
		//Tab2 - Accept

		driver.findElement(By.xpath("//a[@id=\"droppableExample-tab-accept\"]")).click();
		Thread.sleep(5000);
		
		WebElement e1 = driver.findElement(By.xpath("//*[@id=\"acceptable\"]"));
		WebElement e3 = driver.findElement(By.xpath("//*[@id=\"notAcceptable\"]"));
		WebElement e2 = driver.findElement(By.xpath("//div[@id='droppable'][@class='drop-box ui-droppable']"));
		act.clickAndHold(e1).moveToElement(e2).release().build().perform();
		act.clickAndHold(e3).moveToElement(e2).release().build().perform();
		
		//Tab3 - Prevent Propagation
		
		driver.findElement(By.xpath("//a[@id=\"droppableExample-tab-preventPropogation\"]")).click();
		act.clickAndHold(driver.findElement(By.xpath("//div[@id=\"dragBox\"]"))).moveToElement(driver.findElement(By.xpath("//div[@id=\"notGreedyInnerDropBox\"]"))).build().perform();
		WebElement d1 = driver.findElement(By.xpath("//div[@id=\"dragBox\"]"));
		WebElement d2 = driver.findElement(By.xpath("//*[@id=\"notGreedyDropBox\"]/p"));
		act.clickAndHold(d1).moveToElement(d2).build().perform();
		driver.navigate().refresh();
		driver.findElement(By.xpath("//a[@id=\"droppableExample-tab-preventPropogation\"]")).click();
		WebElement x = driver.findElement(By.xpath("//div[@id='dragBox'][@class='drag-box mt-4 ui-draggable ui-draggable-handle']"));
		WebElement d3 = driver.findElement(By.xpath("//div[@id='greedyDropBoxInner'][@class='drop-box ui-droppable']"));
		act.clickAndHold(x).moveToElement(d3).release().build().perform();
		WebElement y = driver.findElement(By.xpath("//div[@id='greedyDropBox'][@class='drop-box-outer mt-4 ui-droppable']"));
		act.clickAndHold(x).moveToElement(y).release().build().perform();
		WebElement z = driver.findElement(By.xpath("//div[@id='greedyDropBox'][@class='drop-box-outer mt-4 ui-droppable']"));		
		act.clickAndHold(x).moveToElement(z).release().build().perform();
		
		//Tab4 - Revert Draggable
		
		driver.get("https://demoqa.com/droppable");
		driver.manage().window().maximize();
		Actions act = new Actions(driver);
		driver.findElement(By.xpath("//a[@id=\"droppableExample-tab-revertable\"]")).click();
		WebElement wr = driver.findElement(By.xpath("//div[@id='revertable'][@class='drag-box mt-4 ui-draggable ui-draggable-handle']"));
		WebElement drp = driver.findElement(By.xpath("//div[@id='revertableDropContainer']//div[@id='droppable']"));
		WebElement nr = driver.findElement(By.xpath("//div[@id='notRevertable'][@class='drag-box mt-4 ui-draggable ui-draggable-handle']"));
		Thread.sleep(5000);
	
		for (int i = 0; i < 15; i++) {
		
		act.dragAndDrop(wr, drp).release().build().perform();
		act.moveByOffset(1, 2);
	}
	
	act.dragAndDrop(nr, drp).release().build().perform();
	
		
		
		//4) https://demo.guru99.com/test/drag_drop.html
		
		driver.get("https://demo.guru99.com/test/drag_drop.html");
		driver.manage().window().maximize();
		Actions action = new Actions(driver);
		
		action.dragAndDrop(driver.findElement(By.xpath("//*[@id=\"fourth\"]/a")), driver.findElement(By.xpath("//*[@id=\"amt7\"]/li"))).perform();
		action.dragAndDrop(driver.findElement(By.xpath("//*[@id=\"fourth\"]/a")), driver.findElement(By.xpath("//*[@id=\"amt8\"]/li"))).perform();
		action.dragAndDrop(driver.findElement(By.xpath("//*[@id=\"credit2\"]/a")), driver.findElement(By.xpath("//*[@id=\"bank\"]/li"))).perform();
		action.dragAndDrop(driver.findElement(By.xpath("//*[@id=\"credit1\"]/a")), driver.findElement(By.xpath("//*[@id=\"loan\"]/li"))).perform();
		Thread.sleep(3000);
		
		System.out.println("Message displayed - "+driver.findElement(By.xpath("//*[@id=\"equal\"]/a")).getText());
		
		
		//5) https://demoqa.com/automation-practice-form ------ fill this Form
		
		
		driver.get("https://demoqa.com/automation-practice-form");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		
		driver.findElement(By.id("firstName")).sendKeys("DemoQA"); //First Name
		Actions action1 = new Actions(driver);
		
		action1.sendKeys(Keys.TAB).sendKeys("Practice FOrm") //Last Name
		.sendKeys(Keys.TAB).sendKeys("DemoQA@gmail.com").sendKeys(Keys.TAB).build().perform();//Email
		Thread.sleep(5000);
		
		driver.findElement(By.xpath("//*[@id=\"genterWrapper\"]/div[2]/div[1]/label")).click();//Gender
		action1.sendKeys(Keys.TAB).sendKeys("1233454545").sendKeys(Keys.TAB).build().perform();//Mobile
		Thread.sleep(5000);
		driver.findElement(By.xpath("//select[@class=\"react-datepicker__month-select\"]")).sendKeys("May");//DOB-Month
		driver.findElement(By.xpath("//select[@class=\"react-datepicker__year-select\"]")).sendKeys("2000");//DOB- Year
		driver.findElement(By.xpath("//*[@id=\"dateOfBirth\"]/div[2]/div[2]/div/div/div[2]/div[2]/div[5]/div[5]")).click();//DOB-Date
		
		action1.sendKeys(Keys.TAB).sendKeys("En").sendKeys(Keys.ENTER).sendKeys("com").sendKeys(Keys.ENTER).build().perform();//Subjects
		Thread.sleep(5000);
		driver.findElement(By.xpath("//*[@id=\"hobbiesWrapper\"]/div[2]/div[1]/label")).click();//Hobbies
		//Thread.sleep(10000);
		//driver.findElement(By.xpath("//input[@id='uploadPicture'][@class='form-control-file']")).click();
		//Thread.sleep(10000);
		//driver.findElement(By.xpath("//input[@id='uploadPicture'][@class='form-control-file']")).sendKeys("\"C:\\Users\\visit\\OneDrive\\Desktop\\Toolsqa.jpg\"");
		driver.findElement(By.xpath("//*[@id=\"currentAddress\"]")).sendKeys("XYZ");//Current Address
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("document.body.style.zoom='50%'");//zoomout
		Thread.sleep(10000);
		action1.sendKeys(Keys.TAB).sendKeys("NCR").sendKeys(Keys.TAB).sendKeys(Keys.TAB).sendKeys("Delhi").sendKeys(Keys.TAB).sendKeys(Keys.TAB).sendKeys(Keys.ENTER).build().perform();
		
		//WebDriverWait w = new WebDriverWait(driver , Duration.ofSeconds(10));
		//WebElement element = w.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"state\"]/div/div[2]/div/svg")));
		//element.click();
		
		//driver.findElement(By.xpath("//div[@class=\" css-1wa3eu0-placeholder\"]")).sendKeys("NCR");
		//driver.findElement(By.xpath("//*[@id=\"state\"]/div/div[2]")).sendKeys("NCR");
		//driver.findElement(By.xpath("//*[@id=\"city\"]/div/div[2]/div")).click();
		//driver.findElement(By.xpath("//*[@id=\"city\"]/div/div[2]/div")).sendKeys("Delhi");
		
		
		
		
		//3) https://demoqa.com/resizable
		
		driver.get("https://demoqa.com/resizable");
		driver.manage().window().maximize();
		
		act.dragAndDropBy(driver.findElement(By.xpath("//*[@id=\"resizableBoxWithRestriction\"]/span")), 150, 70).perform();
		Thread.sleep(10000);
		act.keyDown(Keys.DOWN).keyDown(Keys.DOWN).keyDown(Keys.DOWN).keyUp(Keys.DOWN).build().perform();
		Thread.sleep(10000);
		act.dragAndDropBy(driver.findElement(By.xpath("//body/div[@id='app']/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/span[1]")), 200,0).release().build().perform();
		
		
		
		
				
			
			
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	
	
	}
	


}
