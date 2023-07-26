import com.applitools.eyes.FileLogger;
import com.applitools.eyes.ProxySettings;
import com.applitools.eyes.selenium.Eyes;
import io.appium.java_client.windows.WindowsDriver;
import org.openqa.selenium.*;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.*;

import java.net.MalformedURLException;
import java.net.URL;

public class ExcelDemo
{
    public WindowsDriver driver = null;
    /* Excel; application path */
    public static String appPath="C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
    public static String appArguments="/e C:\\applitools-projects\\winappdriver\\src\\main\\resources\\test.xlsx";
    public static Eyes eyes;


    /* In ideal setup, the same cap instance can be used. This is only for demo */
    @BeforeTest
    public void testSetUp() throws Exception
    {
        DesiredCapabilities capability = new DesiredCapabilities();

        capability.setCapability("ms:experimental-webdriver", true);
        capability.setCapability("app",appPath);
        capability.setCapability("appArguments", appArguments);
        capability.setCapability("platformName", "Windows");
        capability.setCapability("deviceName", "Windows11 VM");

//        WinDriver.start();

        driver = new WindowsDriver(new URL("http://127.0.0.1:4723/"), capability);

        eyes = new Eyes();
        eyes.setLogHandler(new FileLogger("eyes.log", true, true));
        eyes.setForceFullPageScreenshot(false);
        eyes.setApiKey(System.getenv("APPLITOOLS_API_KEY_DEV"));
        //Proxy Setting if required for SDK to communicate with Applitools.
        //eyes.setProxy(new ProxySettings("http://proxyserver", 8080));
        eyes.open(driver, "winappdriver", "Excel Test");
    }

    /* Certain tests are derived from the one present on official WinAppDriver website */
    @Test(description="Demonstration of entering content in Microsoft Excel with Visual Validation", priority = 0)
    public void test_excel() throws InterruptedException, MalformedURLException
    {
        eyes.checkWindow("Blank Sheet");
        driver.findElement(By.name("B2")).click();
        driver.findElement(By.name("Formula Bar")).sendKeys("10");
        driver.findElement(By.name("Formula Bar")).sendKeys(Keys.ENTER);
        eyes.checkWindow("Enter data B2");
        driver.findElement(By.name("C2")).click();
        driver.findElement(By.name("Formula Bar")).sendKeys("20");
        driver.findElement(By.name("Formula Bar")).sendKeys(Keys.ENTER);
        eyes.checkWindow("Enter data C2");
        driver.findElement(By.name("D2")).click();
        driver.findElement(By.name("Formula Bar")).sendKeys("=B2+C2");
        driver.findElement(By.name("Formula Bar")).sendKeys(Keys.ENTER);
        eyes.checkWindow("Confirm Addition Result");
        eyes.close();
    }


    @AfterTest
    public void tearDown()
    {
        if (driver != null)
        {
            /* Instantiated WinAppDriver can be stopped, needs to be started again from terminal */
            driver.close();
            driver.findElement(By.name("Don't Save")).click();
            driver.quit();
            //WinDriver.stop();
        }
    }
}
