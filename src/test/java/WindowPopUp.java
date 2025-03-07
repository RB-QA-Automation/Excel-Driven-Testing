import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class WindowPopUp {

	public static void main(String[] args) {

		WebDriver driver = new ChromeDriver();

		// handling websites which contains pop upon launch

		// when loading URL we can pass in username and password
		// selenium can only test the web application

		driver.manage().window().maximize();

		// load in URL and pass in credentials
		// admin(username):admin(password)
		driver.get("http://admin:admin@the-internet.herokuapp.com/");

		// click on link to trigger pop up
		driver.findElement(By.linkText("Basic Auth")).click();

	}

}
