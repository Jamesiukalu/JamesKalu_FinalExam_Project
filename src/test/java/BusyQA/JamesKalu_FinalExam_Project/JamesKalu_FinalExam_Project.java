package BusyQA.JamesKalu_FinalExam_Project;

import org.testng.annotations.Test;

import org.testng.annotations.BeforeTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.Duration;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;


public class JamesKalu_FinalExam_Project {
//	static Logger logger = Logger.getLogger(LoggerFile.class);
	public static final Logger logger = Logger.getLogger("myClassName");
	
	public static ExtentSparkReporter sparkReporter; //This class is responsible for generating an HTML report with a user-friendly interface. The SparkReporter specifically generates a spark-themed HTML report.
	public static ExtentReports extent;
	public static ExtentTest test;
	
    WebDriver driver;

	public void initializer() {
		sparkReporter =  new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
		sparkReporter.config().setDocumentTitle("Automation Report");
		sparkReporter.config().setReportName("Test Execution Report");
		sparkReporter.config().setTheme(Theme.STANDARD);
		sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
		
		extent = new ExtentReports();
		extent.attachReporter(sparkReporter);		
	}
	
	public static String captureScreenshot(WebDriver driver) throws IOException {
		String FileSeparator = System.getProperty("file.separator"); // "/" or "\"
		String Extent_report_path = "."+FileSeparator+"Reports"; // . means parent directory
		File Src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		String Screenshotname = "screenshot"+Math.random()+".png";
		File Dst = new File(Extent_report_path+FileSeparator+"Screenshots"+FileSeparator+Screenshotname);
		FileUtils.copyFile(Src, Dst);
		String absPath = Dst.getAbsolutePath();
		System.out.println("Absolute path is:"+absPath);
		return absPath;
	}
	
    @Test
    public void newTest() throws InterruptedException {
		String methodName = new Exception().getStackTrace()[0].getMethodName();
		String className = new Exception().getStackTrace()[0].getClassName();
		test = extent.createTest(methodName,"FinalExam_Project Testing");
		test.log(Status.INFO, "Two(2) Screen prints captured for each and every record in 5 tables");
		test.assignCategory("Regression Testing");
		
  	  String excelFilePath = "C:\\Users\\BUYPC COMPUTERS\\SeleniumAssignmentDocs\\excelFile.xlsx";
  	  
  	  try {

  		  FileOutputStream outputStream = new FileOutputStream(excelFilePath);  // Write changes to output Excel file
          //open workbook
          XSSFWorkbook workbook = new XSSFWorkbook();
   
          Thread.sleep(1000);

          
          // Wait for the page to load and get the tables
//          WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
//          wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("BILLETS_heading")));
          
       // Locate the table section
          int count = 0;
          while (count < 5) { //process only the first 5 tables
              List<WebElement> tableRows = driver.findElements(By.xpath("//table[contains(@class, 't-Report-report')]/tbody/tr"));

           // Loop through each row
              for (WebElement row : tableRows) {
                  // Locate the first cell in the current row
                  List<WebElement> nameLinks = row.findElements(By.xpath(".//td[1]/a")); // Use relative XPath to find <a> in the first <td>

                  // Check if the first <td> contains a link
                  if (!nameLinks.isEmpty()) {  // Ensure the list is not empty
                      WebElement nameLink = nameLinks.get(0); // Get the first link
              
                      // Check if the nameLink is displayed and is an anchor element
                      if (nameLink.isDisplayed() && nameLink.getTagName().equals("a")) {
                          String name = nameLink.getText(); // Get the text of the link
                          // Wait for the link to be clickable
                   //       wait.until(ExpectedConditions.elementToBeClickable(nameLink));
                          Thread.sleep(1000); // Sleep for 1 second
                          test.addScreenCaptureFromPath(captureScreenshot(driver));
                          nameLink.click(); // Click on the name link
                  
                          Thread.sleep(2000); 
        
                //      wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table[@width='500' and @border='1']"))); // Wait for small table

                      // Copy the small table from the new page
                      List<WebElement> smallTableRows = driver.findElements(By.xpath(".//tbody/tr"));

                      // Create a new Excel sheet with the same name as the link
                    XSSFSheet sheet = workbook.createSheet(name);
                 // Fill the Excel sheet with data from the small table
                    int rowIndex = 0;
                    for (WebElement smallRow : smallTableRows) {
                  	  
                        XSSFRow excelRow = sheet.createRow(rowIndex++);
                        List<WebElement> cells = smallRow.findElements(By.tagName("td"));
                        for (int i = 0; i < cells.size(); i++) {
                            XSSFCell cell = excelRow.createCell(i);
                            cell.setCellValue(cells.get(i).getText());
                        }
                    }
                    // Close the modal
                    WebElement closeButton = driver.findElement(By.xpath("//button[@title='Fermer']"));
                    test.addScreenCaptureFromPath(captureScreenshot(driver));
                    closeButton.click();
               //     wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("BILLETS_heading"))); // Wait for the main page to reload
                    Thread.sleep(1000); 
                    
                    // Increment the count of processed tables
                    count++;
                    if (count >= 5) break; // Process only the first 5 tables
            	}
              }

          }
         }
          workbook.write(outputStream); // Write to the Excel file
       	 outputStream.close();  // Close outputStream
         workbook.close();
         
		  logger.info("This test is successful");
 
      } catch (IOException e) {
          e.printStackTrace();
      }

    }
  	  


    @BeforeTest
    public void beforeTest() {
		initializer();
        driver = new ChromeDriver();
        driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
    //    driver.manage().window().maximize();
	//	driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		  logger.info("This test started successful");
    }

    @AfterTest
    public void afterTest() {
		extent.flush();
		driver.close();
        driver.quit();
    }
}
