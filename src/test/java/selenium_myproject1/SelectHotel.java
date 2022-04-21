package selenium_myproject1;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class SelectHotel extends BaseClass {
	public SelectHotel() {
	    PageFactory.initElements(driver, this);
	    }
		
		@FindBy(id="radiobutton_0")
		private WebElement select;
		
		@FindBy(id= "continue")
		private WebElement Continue;

		public WebElement getSelect() {
			return select;
		}

		public WebElement getContinue() {
			return Continue;
		}
		

}
