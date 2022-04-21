package selenium_myproject1;
import java.util.Date;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class junit {
	
		@BeforeClass
		public static void beforeclass() {
			System.out.println("Browser configuration and launch");

		}
		@AfterClass
	    public static void afterclass() {
			System.out.println("close browser");
		

		}
		
		@Before
		public void beforeeachmethod() {
			Date d = new Date();
	     System.out.println("Start time:"+d);
		}

		@After
		public void aftereachmethod() {
			Date d = new Date();
			System.out.println("End Time:"+d);

		}
		@Test
		public void Tc1() {
			System.out.println("TC1.....");
		}
		@Test
		public void Tc2() {
			System.out.println("TC2....");

		}
		@Test
		public void TC3() {
			System.out.println("TC3.....");

		}
		
	}


