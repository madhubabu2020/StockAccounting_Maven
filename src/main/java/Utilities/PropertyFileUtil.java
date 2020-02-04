package Utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

public class PropertyFileUtil {
	
	
	public static String getValueForKey(String Key) throws Exception{
		
		Properties configProperties=new Properties();
		
		FileInputStream fis=new FileInputStream("D:\\madhu.p\\StockAccounting_Maven\\PropertiesFile\\Envoirnment.properties");
		
		configProperties.load(fis);
		
		return configProperties.getProperty(Key);
		
	}

}

	