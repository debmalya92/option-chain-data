package com.config;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Run_Configuration {
	
	public WebDriver driver;

	
	@BeforeClass
	public void setup() {
		
		DesiredCapabilities capabilities = DesiredCapabilities.chrome();
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--headless");
		capabilities.setCapability(ChromeOptions.CAPABILITY, options);
		System.setProperty("webdriver.chrome.driver","resources\\chromedriver.exe");	
		driver = new ChromeDriver(options); 
		driver.get("https://web.sensibull.com/home");
		driver.manage().window().setSize(new Dimension(1280, 720));
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		
	}
	
	public void navigateToOptionChain() throws InterruptedException {

		System.out.println("Executing component - navigateToOptionChain");
		WebDriverWait wait = new WebDriverWait(driver, 10);

		//Click on Analyze option
		By btn_analyze = By.xpath("//*[contains(text(),'Analyze')]/parent::button");
		wait.until(ExpectedConditions.elementToBeClickable(btn_analyze));
		driver.findElement(btn_analyze).click();
		System.out.println("[ Analyze ] clicked");

		//Click on Option Change option
		By btn_option_chain = By.xpath("//a[@href='/option-chain']");
		//wait.until(ExpectedConditions.elementToBeClickable(btn_option_change));
		driver.findElements(btn_option_chain).get(1).click();
		Thread.sleep(2000);

		By hdr_option_chain = By.xpath("//a[@href='/home']/following-sibling::h1");
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(btn_analyze));
		if(driver.findElement(hdr_option_chain).getAttribute("textContent").equals("Option Chain")) {
			System.out.println("[ Option Chain ] clicked");
		}

	}
		
	
	public void setOptionChainPage() throws IOException, InterruptedException {
		System.out.println("Executing component - setOptionChainPage");
		WebDriverWait wait = new WebDriverWait(driver, 15);
		
		By settings = By.xpath("//div[contains(@class, 'FiltersRightWrapper')]/button");
		By oi_change = By.xpath("//*[@label='OI Change']");
		By lbl_oi_change = By.xpath("//*[@id='oc-table-head']/div/div/div");
		
		//Adding OI Change value 
		wait.until(ExpectedConditions.visibilityOfElementLocated(settings));
		driver.findElement(settings).click();
		System.out.println("[ Settings ] clicked");
		
		wait.until(ExpectedConditions.elementToBeClickable(oi_change));
		WebElement elmnt = driver.findElement(oi_change);
		if(!elmnt.getAttribute("class").contains("Mui-checked")) { 
			elmnt.click();
			driver.findElement(By.xpath("//span[text()='Close']")).click();
		}
		
		//OI Change column added validation
		wait.until(ExpectedConditions.visibilityOfElementLocated(lbl_oi_change));
		if(driver.findElement(lbl_oi_change).getAttribute("textContent").equals("OIchange")) {
			System.out.println("OI Change Value column added"); 
		}
		Thread.sleep(2000);
		
		//handling "See Trades" popup
		By btn_seeTrade_close = By.xpath("//div[@class='summary']/following-sibling::div/button");
		List<WebElement> element = driver.findElements(btn_seeTrade_close);
		if(!element.isEmpty()) {
			element.get(0).click();
			System.out.println("See Trade popup closed"); 
		}
		
	}
	
	public String getDate() {
		String s_date = "";
		SimpleDateFormat d_formatter = new SimpleDateFormat("dd-MM-yyyy");
		s_date = d_formatter.format(new Date());
		return s_date;
	}
	
	public String getTime() {
		String s_time="";
		SimpleDateFormat t_formatter = new SimpleDateFormat("hh-mm aa");
		s_time = t_formatter.format(new Date());
		return s_time;
	}
	public void getScreenshot() throws IOException {
		System.out.println("Executing component - getScreenshot");
		File file_image = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		String curr_date = getDate();
		String curr_time = getTime();
		
	    
	    File f = new File(curr_date);
	    if(!f.exists()) {
	    	f.mkdir();
	    	System.out.println("Current Date Folder created");
	    }else {
	    	System.out.println("Folder already exist");
	    }
	    
	    FileUtils.copyFile(file_image, new File(curr_date+"\\"+curr_time+".jpg"));
	    System.out.println("Screenshot added to today's folder");
	}
	
	
	public void storeScreenshotToWord() throws IOException, InvalidFormatException {
		System.out.println("Executing component - storeScreenshotToWord");
		int width = 1080;
		int height = 720;
		//create blank doc & create paragraph
		XWPFDocument document = new XWPFDocument();
		XWPFRun run = document.createParagraph().createRun();

		// Step 3: Creating a File output stream of word
		File doc_file = new File(getDate()+".docx");
		if(doc_file.exists()) {
			doc_file.delete();
		}
		doc_file.createNewFile();
		FileOutputStream fout = new FileOutputStream(doc_file, true);


		File file = new File(getDate());
		File[] images = file.listFiles(new FilenameFilter() {
			
			@Override
			public boolean accept(File dir, String name) {
				
				return name.endsWith(".jpg");
			}
			
		});
		
		for(File im: images) {
			FileInputStream fin = new FileInputStream(im);
			
			String str_time = FilenameUtils.removeExtension(im.getName()).replace("-", ":");
			run.setBold(true);
			run.setFontSize(16);
			run.setText(str_time.toUpperCase());
			
			//Adding the picture in the doc
			run.addPicture(fin, XWPFDocument.PICTURE_TYPE_JPEG, im.getName(),Units.toEMU(width), Units.toEMU(height));
			
		}
		document.write(fout);
		System.out.println("Screenshot added to doc file");

		fout.close();
		document.close();
	}

}
