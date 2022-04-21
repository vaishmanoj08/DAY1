package selenium_myproject1;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class BookHotel extends BaseClass {

	public BookHotel() {
		PageFactory.initElements(driver,this);
	} 
	@FindBy(id="first_name")
	private WebElement First_name;
	
	@FindBy(id="last_name")
	private WebElement Last_name;
	
	@FindBy(id="address")
	private WebElement Billing_Address;
	
	@FindBy(id="cc_num")
	private WebElement cc_num;
	
	@FindBy(id="cc_type")
	private WebElement cctype;
	
	@FindBy(id="cc_exp_month")
	private WebElement cc_exp_month;
	
	@FindBy(id="cc_exp_year")
	private WebElement expiryyear;
	
	
	public WebElement getExpiryyear() {
		return expiryyear;
	}
	@FindBy(id="cc_cvv")
	private WebElement cc_cvv;
	
	@FindBy(id="book_now")
	private WebElement book_now;

	public WebElement getFirst_name() {
		return First_name;
	}

	public WebElement getLast_name() {
		return Last_name;
	}

	public WebElement getBilling_Address() {
		return Billing_Address;
	}

	public WebElement getCc_num() {
		return cc_num;
	}

	public WebElement getCctype() {
		return cctype;
	}

	public WebElement getCc_exp_month() {
		return cc_exp_month;
	}

	public WebElement getCc_cvv() {
		return cc_cvv;
	}

	public WebElement getBook_now() {
		return book_now;
	}

}
