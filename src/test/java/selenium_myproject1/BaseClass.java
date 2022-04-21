package selenium_myproject1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	static WebDriver driver;
    Actions actions;
    static Select select;
    Alert alert;
    TakesScreenshot screenshot;
	CellType celltype;
    
	public static void BrowserConfigChrome() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		
	}
	
	public static void openurl(String url) {
		driver.get(url);

	} 
	
		
   public static void maximizewindow() {
driver.manage().window().maximize();
}
    
    
    public static WebElement findLocatorsbyid(String id) {
    WebElement findElement = driver.findElement(By.id(id));
    return findElement;
    
	}
    
   public static WebElement findLocatorsbyname(String name) {
   WebElement findElement = driver.findElement(By.name(name));
   return findElement;	
	
}

   public static void InputText(WebElement element, String value) {
   element.sendKeys(value);

}
   public static void click(WebElement element) {
   element.click();
}
   public void Closebrowser() {
   driver.close();

}
  public void quitpage() {
	driver.quit();
}
  public void refreshpage() {
	  driver.navigate().refresh();

}
  public void doubleclick(WebElement element) {
	Actions actions = new Actions(driver);
	actions.doubleClick(element).perform();
}
  public void rightclick(WebElement element) {
	actions.contextClick(element).perform();

}
  public void movetoelement(WebElement element) {
	actions.moveToElement(element).perform();

}
  public String swithchtoparentwindow(WebElement element, String value ) {
	String windowHandle = driver.getWindowHandle();
	return windowHandle;

}
 public static void dropdownselectbyvalue(WebElement element, String value) {
	Select select = new Select(element);
	select.selectByValue(value);
}
 
 public static void dropdownselectbyindex(WebElement element, int index) {
	 Select select = new Select(element);
	 select.selectByIndex(index);
	}
 public static void ddnvisibletext(WebElement element, String text) {
		
	 Select select =new Select(element);
		select.selectByVisibleText(text);
	}
 public static void dropdowndeselectbyvalue(WebElement element, String value) {
		Select select = new Select(element);
		select.deselectByValue(value);
	}
	 
	 public static void dropdowndeselectbyindex(WebElement element, int index) {
			Select select = new Select(element);	
		    select.deselectByIndex(index);
		}
	 public static void dropdowndeselectbytext(WebElement element, String text) {
			Select select = new Select(element);	
		 select.deselectByVisibleText(text);
		}
	 
   public void deselectall() {
	   select.deselectAll();
	   
}
   public void ismultiple() {
       select.isMultiple();
}
   public void getalloptions() {
	List<WebElement> options = select.getOptions();
	for (int i=0; i>options.size();i++);
	{
		WebElement element = options.get(0);
        String string = element.getText();	
	}

}
   public void forwardpage() {
   driver.navigate().forward();
}
   public void backwardpage() {
   driver.navigate().back();
}
   public void simplealert() {
   Alert alert2 = driver.switchTo().alert();
   alert2.accept();
   
   
}
   public void confirmalert() {
   Alert alert2 = driver.switchTo().alert();
   alert2.dismiss();
  }
  public void draganddrop(WebElement source, WebElement target) {
	actions.dragAndDrop(source, target);

  } 
 public String excelread(String path, String sheet, int rowindex, int cellindex) throws IOException {
   String value = null;
   File F= new File(path);
   FileInputStream fin = new FileInputStream(F);
   Workbook W =new XSSFWorkbook(fin);
   Sheet s = W.getSheet(sheet);
   Row r = s.getRow(rowindex);
   Cell c = r.getCell(cellindex);
 
   
   switch (celltype) {
   case STRING:
   value =c.getStringCellValue();
   break;
   
   
case NUMERIC:
	if(DateUtil.isCellDateFormatted(c))
	{
		Date datecellvalue = c.getDateCellValue();
	    SimpleDateFormat sim = new SimpleDateFormat("dd-mm-yy");
	    		value=sim.format(datecellvalue);
	    		System.out.println("value");
	}else {
		Date numericcellvalue= c.getDateCellValue();
		BigDecimal big = BigDecimal.valueOf(cellindex);
		value =big.toString();
		System.out.println(value);
	}
	break;

default:
	System.out.println("invalid");
	break;
   }
return value;
 
   }
 


 public String excelwrite(String path, String sheet, int rowindex, int cellindex) throws IOException {
	   String value = null;
	   File F= new File(path);
	   FileInputStream fin = new FileInputStream(F);
	   Workbook W =new XSSFWorkbook(fin);
	   Sheet s = W.getSheet(sheet);
	   Row r = s.getRow(rowindex);
	   Cell c = r.getCell(cellindex);
	   c.setCellValue(value);
	   FileOutputStream fout = new FileOutputStream(F);
	   
	   W.write(fout);
	return value;
 }
	   
	   public static String getthetext(WebElement element) {
	   String text = element.getText();
	   return text;
	   }
	   public static String getattribute(WebElement element, String value) {
       String attribute = element.getAttribute(value);
       return attribute;
	   }
}


