package com.testcases;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

import com.config.Run_Configuration;

public class SensiBullOIData extends Run_Configuration {
	
	@Test
	public void GetSensiBullOIData() throws InvalidFormatException, IOException, InterruptedException {
		navigateToOptionChain();
		setOptionChainPage();
		getScreenshot();
		storeScreenshotToWord();
	}

}
