package demo;

import java.io.File;
import java.io.FileInputStream;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONObject;

import com.opensymphony.xwork2.ActionContext;
import com.opensymphony.xwork2.ActionSupport;

public class ExcelUpLoadAction extends ActionSupport {
	private File file;
	private String jsonDatas;
	private String colName;

	public String getJsonDatas() {
		return jsonDatas;
	}

	public void setJsonDatas(String jsonDatas) {
		this.jsonDatas = jsonDatas;
	}

	public String getColsName() {
		return colName;
	}

	public void setColsName(String colsName) {
		this.colName = colsName;
	}

	public ExcelUpLoadAction() {
		super();
	}

	// 解析excel文件

	public List<Map> readExcelDatas(String[] colsName) {
		// 获取上传Excel文件的文件流
		List<Map> list = new ArrayList<Map>();
		try {
			FileInputStream in = new FileInputStream(file);
			Workbook wb = WorkbookFactory.create(in);
			Sheet sheet1 = wb.getSheetAt(0);
			for (Row row1 : sheet1) { // 遍历所有行
				Map<String, String> map = new HashMap<String, String>();
				boolean flag = false;
				for (Cell cell1 : row1) { // 遍历行内的单元格
					if (cell1.getColumnIndex() <= colsName.length) {
						if (!URLDecoder.decode(cell1.toString(), "utf-8")
								.equals("")) {
							map.put(colsName[cell1.getColumnIndex()],
									URLDecoder.decode(cell1.toString(), "utf-8"));
						}
						if (flag != true) {
							if (!cell1.toString().equals("")) {
								flag = true;
							}
						}
					}
				}
				if (flag) {
					list.add(map);
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString());
			System.out.println("Excel文件获取失败！");
		}
		return list;
	}

	public String excelUpLoad() throws Exception {
		List<Map> resultList = new ArrayList<Map>();
		// excel表头列名
		String colsName[] = { "name", "gender", "age", "idNo", "phoneNumber" };
		List<Map> list = new ArrayList<Map>();
		list = readExcelDatas(colsName);
		boolean flage = true;
		colName = "name_gender_age_idNo_phoneNumber";
		for (Map map : list) {
			if (flage) {
				// 对表头进行处理
				flage = false;
				resultList.add(map);
				continue;
			}
			if (valIsNull(map, colName)) {		
				resultList.add(map);
				continue;
			}
			resultList.add(map);
			
		}
		JSONObject json = new JSONObject();
		json.put("row", resultList);
		jsonDatas = json.toString();
		ActionContext context = ActionContext.getContext();
		context.put("colName", colName);
		Map request = (Map) context.getSession();
		// //返回excel解析后的数据
		// request.put("jsonDatas", jsonDatas.toString());
		// //返回前台页面表头的列名
		request.put("colName", colName);
		return SUCCESS;
	}

	// 判断值是否为空
	private boolean valIsNull(Map map, String key) {
		String[] keyArr = key.split("_");
		for (int i = 0; i < keyArr.length; i++) {
			if (map.get(keyArr[i]) != null
					&& map.get(keyArr[i]).toString().trim().equals("")) {
				return true;
			}
		}
		return false;
	}

	public void setFile(File file) {
		this.file = file;
	}

	public File getFile() {
		return file;
	}
}
