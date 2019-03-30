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
	
	private WebDriver driver = null; // 驱动器
	private String baseUrl; // 访问的网址
	
	// 取用户名的后六位作为密码，不足六位则返回空串
	private String getPassword(String id){
		String ret = "";
		if(id.length() >= 6) {
			ret = id.substring(id.length() - 6);
		}
		return ret;
	}
	
	@Before
	public void setUp() throws Exception {
		String driverPath = "D:/Program Files/Mozilla Firefox/geckodriver.exe"; // 浏览器驱动器的位置
		System.setProperty("webdriver.firefox.bin","D:/Program Files/Mozilla Firefox/firefox.exe"); // 浏览器非默认安装位置，需要设置浏览器位置
		System.setProperty("webdriver.gecko.driver", driverPath); // 设置浏览器驱动器
		driver = new FirefoxDriver(); // 生成 Firefox 浏览器的驱动器
		//String driverPath = "C:/Program Files (x86)/Google/Chrome/Application/chromedriver.exe";
		//System.setProperty("webdriver.chrome.driver", driverPath);
		//driver = new ChromeDriver();
		baseUrl = "http://121.193.130.195:8800/login"; // 测试网址
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); // 全局设置隐式等待30秒
	}
	
	// 读取Excel文件中的数据
	private Collection<String[]> getData() {
		File file = new File("test/cn/tjuscsst/lab2/软件测试名单.xlsx"); // 数据文件
		ArrayList<String[]> list = new ArrayList<String[]>(); // 将数据放到列表中
		try {
			FileInputStream fis = new FileInputStream(file);		
			XSSFWorkbook xwb = new XSSFWorkbook(fis);
			Sheet sheet = xwb.getSheetAt(0); // 取第一个表
			
			int firstRowNum = sheet.getFirstRowNum(); // 首行号
			int lastRowNum = sheet.getLastRowNum(); // 尾行号
			
			Row row = null;
			Cell cell_a = null;
			Cell cell_b = null;
			Cell cell_c = null;
			for(int i = firstRowNum + 2; i <= lastRowNum; i++) {
				row = sheet.getRow(i); // 取得第三行，从第三行开始取，因为第二行是表头
				cell_a = row.getCell(1); // 取学号列
				BigDecimal bd = new BigDecimal(cell_a.getNumericCellValue());
				String cell_a_Value = bd.toString();
				
				cell_b = row.getCell(2); // 取姓名列
				String cell_b_Value = cell_b.getStringCellValue().trim();
				
				cell_c = row.getCell(3); // 取git地址列
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
		driver.get(baseUrl); // 登入网址
		Collection<String[]> alist = this.getData(); // 获取数据
		for(String[] strings : alist){
			driver.findElement(By.name("id")).clear(); // 清空ID文本框
			driver.findElement(By.name("id")).sendKeys(strings[0]); // 输入ID
			driver.findElement(By.name("password")).clear(); // 清空密码文本框
			driver.findElement(By.name("password")).sendKeys(getPassword(strings[0])); // 输入密码
			driver.findElement(By.id("btn_login")).sendKeys(Keys.ENTER); // 点击登录
			// 对比用户信息
			assertEquals(strings[1], 
					driver.findElement(By.id("student-id")).getText() + "-" +
					driver.findElement(By.id("student-name")).getText() + "-" +
					driver.findElement(By.id("student-git")).getText());
			driver.findElement(By.id("btn_logout")).sendKeys(Keys.ENTER); // 用户登出
			driver.findElement(By.id("btn_return")).sendKeys(Keys.ENTER); // 返回主页
			driver.manage().deleteAllCookies(); // 删除缓存
		}
		
	}

	@After
	public void tearDown() throws Exception {
		//driver.close();
		driver.quit(); // 退出浏览器
	}
}