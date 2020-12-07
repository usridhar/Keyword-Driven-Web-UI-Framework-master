package com.prac.tests;

import org.testng.Reporter;
import org.testng.annotations.Test;

import com.prac.keyword.engine.KeyWordEngine;
/**
 * 
 * @author Umarani Sridhar
 *
 */
public class AutoQuoteTest {
	
	public KeyWordEngine keyWordEngine;
	
	@Test(priority=1)
	public void autoQuoteTest() {
		keyWordEngine = new KeyWordEngine();
		//keyWordEngine.startExecution("AutoQuote");
		keyWordEngine.mainExecution();
		Reporter.log("Generate Auto Quote");
	}
	
	

}
