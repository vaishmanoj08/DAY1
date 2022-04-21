package selenium_myproject1;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.CacheLookup;
import org.openqa.selenium.support.FindAll;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.FindBys;
import org.openqa.selenium.support.PageFactory;

import com.github.dockerjava.api.model.Driver;

public class Login extends BaseClass {
	public Login() {
		PageFactory.initElements(driver, this);
	}


	@FindBy(id="username")
	
	private WebElement username;

	@FindBy(id="password")
       
	private WebElement password;
	
	@FindBy(name="login")
	private WebElement login;
	public WebElement getUsername() {
		return username;
	}
	public WebElement getPassword() {
		return password;
	}
	public WebElement getLogin() {
		return login;
	}
	
	
}


