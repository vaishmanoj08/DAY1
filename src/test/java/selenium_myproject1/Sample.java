package selenium_myproject1;



import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;
public class Sample {

public static void main(String[] args) {

	
WebDriverManager.chromedriver().setup();
WebDriver driver = new ChromeDriver();
driver.get("https://en-gb.facebook.com/");
WebElement txtusername = driver.findElement(By.id("email"));
txtusername.sendKeys("vaishu");
WebElement txtpass = driver.findElement(By.id("pass"));
txtpass.sendKeys("1234");
WebElement btnlogin = driver.findElement(By.name("login"));
btnlogin.click();

	}

}



