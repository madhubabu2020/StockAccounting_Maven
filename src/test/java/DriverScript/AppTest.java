package DriverScript;

import org.testng.annotations.Test;

public class AppTest {
  @Test
  public void starttest() throws Throwable {
	  DriverFactory dr= new DriverFactory();
	  dr.startTest();
  }
}
