package selenium_myproject1;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class SearchHotel extends BaseClass{
	public SearchHotel() {
		PageFactory.initElements(driver, this);;
	}
	
	@FindBy(id="location")		 
      private WebElement Locations;	
      
     @FindBy(id="hotels")
	  private WebElement Hotels;
		 
	  @FindBy(id="room_type")
	  private WebElement room_type;
	 
	  @FindBy(id="room_nos")
	  private WebElement room_nos;
	  
	  @FindBy(id="datepick_in")
	  private WebElement datepick_in;
	  
	  @FindBy(id="datepick_out")
	  private WebElement datepick_out;
	  
	  @FindBy(id="adult_room")
	  private WebElement adult_room;
	  
	  @FindBy(id="child_room")
	  private WebElement child_room;
	  
	  @FindBy(id="Submit")
	  private WebElement Search;

	public WebElement getLocations() {
		return Locations;
	}

	public WebElement getHotels() {
		return Hotels;
	}

	public WebElement getRoom_type() {
		return room_type;
	}

	public WebElement getRoom_nos() {
		return room_nos;
	}

	public WebElement getDatepick_in() {
		return datepick_in;
	}

	public WebElement getDatepick_out() {
		return datepick_out;
	}

	public WebElement getAdult_room() {
		return adult_room;
	}

	public WebElement getChild_room() {
		return child_room;
	}

	public WebElement getSearch() {
		return Search;
	}

	

	
		
	}



