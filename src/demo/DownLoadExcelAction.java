package demo;

import java.io.FileOutputStream;
import java.io.InputStream;

import javax.servlet.ServletContext;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.struts2.ServletActionContext;
import org.json.JSONArray;
import org.json.JSONObject;

import com.opensymphony.xwork2.ActionContext;
import com.opensymphony.xwork2.ActionSupport;

public class DownLoadExcelAction extends ActionSupport{
	private String data;
	private String cols;
	private String fileName;
	

	


	public InputStream getDownloadFile() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet =  wb.createSheet("new sheet");
		JSONObject json = new JSONObject(data);
		JSONArray jsonRow = json.getJSONArray("row");
		String []col = cols.split("_"); 
		for(int index=0; index<jsonRow.length(); index++){
			//�½�excelһ�б��
			Row row = sheet.createRow((short)(index));
			JSONObject tmp = (JSONObject)jsonRow.get(index);
			for(int colIndex=0; colIndex<col.length; colIndex++){
				//�����ݲ���excel��Ԫ����
				row.createCell(colIndex).setCellValue(tmp.getString(col[colIndex]));
			}
		}
		ActionContext context = ActionContext.getContext();
		// ���ServletContext����
		ServletContext servletContext = (ServletContext) context.get(ServletActionContext.SERVLET_CONTEXT);
		//��ȡ��Ŀ��Ŀ¼��ַ
		String path = servletContext.getRealPath("") + "\\";
		System.out.println(path);
		this.fileName = "demo_downLoad.xls";
		FileOutputStream fileOut = new FileOutputStream(path+fileName);
		//�����ɵ�excel����д���ļ�
		wb.write(fileOut);
		fileOut.close();
		return  ServletActionContext.getServletContext().getResourceAsStream(fileName); 
	}
	

	public String execute() throws Exception {   
		  
        return SUCCESS;   
  
    }
	public String getCols() {
		return cols;
	}



	public void setCols(String cols) {
		this.cols = cols;
	}



	public String getData() {
		return data;
	}

	public void setData(String data) {
		this.data = data;
	}
	
	public String getFileName() {
		return fileName;
	}


	public void setFileName(String fileName) {
		this.fileName = fileName;
	}
}
