package cn.tjuscsst.lab2;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.*;
import static org.junit.Assert.*;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collection;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;

public class TestSelenium {
	
	private WebDriver driver = null; // ������
	private String baseUrl; // ���ʵ���ַ
	
	// ȡ�û����ĺ���λ��Ϊ���룬������λ�򷵻ؿմ�
	private String getPassword(String id){
		String ret = "";
		if(id.length() >= 6) {
			ret = id.substring(id.length() - 6);
		}
		return ret;
	}
	
	@Before
	public void setUp() throws Exception {
		String driverPath = "D:/Program Files/Mozilla Firefox/geckodriver.exe"; // �������������λ��
		System.setProperty("webdriver.firefox.bin","D:/Program Files/Mozilla Firefox/firefox.exe"); // �������Ĭ�ϰ�װλ�ã���Ҫ���������λ��
		System.setProperty("webdriver.gecko.driver", driverPath); // ���������������
		driver = new FirefoxDriver(); // ���� Firefox �������������
		//String driverPath = "C:/Program Files (x86)/Google/Chrome/Application/chromedriver.exe";
		//System.setProperty("webdriver.chrome.driver", driverPath);
		//driver = new ChromeDriver();
		baseUrl = "http://121.193.130.195:8800/login"; // ������ַ
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); // ȫ��������ʽ�ȴ�30��
	}
	
	// ��ȡExcel�ļ��е�����
	private Collection<String[]> getData() {
		File file = new File("test/cn/tjuscsst/lab2/�����������.xlsx"); // �����ļ�
		ArrayList<String[]> list = new ArrayList<String[]>(); // �����ݷŵ��б���
		try {
			FileInputStream fis = new FileInputStream(file);		
			XSSFWorkbook xwb = new XSSFWorkbook(fis);
			Sheet sheet = xwb.getSheetAt(0); // ȡ��һ����
			
			int firstRowNum = sheet.getFirstRowNum(); // ���к�
			int lastRowNum = sheet.getLastRowNum(); // β�к�
			
			Row row = null;
			Cell cell_a = null;
			Cell cell_b = null;
			Cell cell_c = null;
			for(int i = firstRowNum + 2; i <= lastRowNum; i++) {
				row = sheet.getRow(i); // ȡ�õ����У��ӵ����п�ʼȡ����Ϊ�ڶ����Ǳ�ͷ
				cell_a = row.getCell(1); // ȡѧ����
				BigDecimal bd = new BigDecimal(cell_a.getNumericCellValue());
				String cell_a_Value = bd.toString();
				
				cell_b = row.getCell(2); // ȡ������
				String cell_b_Value = cell_b.getStringCellValue().trim();
				
				cell_c = row.getCell(3); // ȡgit��ַ��
				String cell_c_Value = cell_c.getStringCellValue().trim();
				
				String result = cell_a_Value + "-" + cell_b_Value + "-" + cell_c_Value;
				String[] obj = {cell_a_Value, result};
				list.add(obj);
			}
		} catch(Exception e) {
			
		}
		return list;
	}

	@Test
	public void testSelenium() throws Exception {
		driver.get(baseUrl); // ������ַ
		Collection<String[]> alist = this.getData(); // ��ȡ����
		for(String[] strings : alist){
			driver.findElement(By.name("id")).clear(); // ���ID�ı���
			driver.findElement(By.name("id")).sendKeys(strings[0]); // ����ID
			driver.findElement(By.name("password")).clear(); // ��������ı���
			driver.findElement(By.name("password")).sendKeys(getPassword(strings[0])); // ��������
			driver.findElement(By.id("btn_login")).sendKeys(Keys.ENTER); // �����¼
			// �Ա��û���Ϣ
			assertEquals(strings[1], 
					driver.findElement(By.id("student-id")).getText() + "-" +
					driver.findElement(By.id("student-name")).getText() + "-" +
					driver.findElement(By.id("student-git")).getText());
			driver.findElement(By.id("btn_logout")).sendKeys(Keys.ENTER); // �û��ǳ�
			driver.findElement(By.id("btn_return")).sendKeys(Keys.ENTER); // ������ҳ
			driver.manage().deleteAllCookies(); // ɾ������
		}
		
	}

	@After
	public void tearDown() throws Exception {
		//driver.close();
		driver.quit(); // �˳������
	}
}