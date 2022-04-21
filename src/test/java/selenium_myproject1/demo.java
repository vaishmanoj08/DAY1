package selenium_myproject1;

import org.openqa.selenium.WebElement;

public class demo extends BaseClass{

	public static void main(String[] args) {

		BrowserConfigChrome();
		openurl("https://adactinhotelapp.com/");
		maximizewindow();
		Login l =new Login();
		InputText(l.getUsername(),"vaishvaishu022");
		InputText(l.getPassword(), "1234567");
	    click(l.getLogin());
		SearchHotel s = new SearchHotel();	

        ddnvisibletext(s.getLocations(), "Sydney");
		ddnvisibletext(s.getHotels(), "Hotel Creek");
		ddnvisibletext(s.getRoom_type(), "Standard");
		dropdownselectbyindex(s.getRoom_nos(), 1);
		dropdownselectbyindex(s.getAdult_room(), 2);
		dropdownselectbyindex(s.getChild_room(), 0);
		click(s.getSearch());
		
		SelectHotel sl= new SelectHotel();
		click(sl.getSelect());
		click(sl.getContinue());
		
       
        BookHotel b = new BookHotel();
        InputText(b.getFirst_name(), "vaishnavi");
        InputText(b.getLast_name(), "sundararajan");
        InputText(b.getBilling_Address(), "chennai");
        InputText(b.getCc_num(), "1234567890123456");
        ddnvisibletext(b.getCctype(), "VISA");
        ddnvisibletext(b.getCc_exp_month(),"January");
        ddnvisibletext(b.getExpiryyear(), "2016");
        InputText(b.getCc_cvv(), "123");
        findLocatorsbyname("book_now");
        click(b.getBook_now());
		
	}
	
}
