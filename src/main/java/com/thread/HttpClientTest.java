package com.thread;

import org.apache.http.Header;
import org.apache.http.HttpEntity;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.*;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

/**
 * httpclient保持登录状态实例,有待优化
 * Created by liliang on 2018/1/23.
 */
public class HttpClientTest {
    private Header[] headers;
    private static String urlLogin = "https://crm.xiaoshouyi.com/global/do-login.action";

    private String post = "https://crm.xiaoshouyi.com/json/crm_expense/save.action";
    //Excel文件路径
    private String PATH = "/Users/liliang/Desktop/xiaoshouyi1.xlsx";

    public XSSFRow row;
    //转换为cookie
//    public CookieStore getCookie(Header[] headers){
//        CookieStore cookieStore =new BasicCookieStore();
//        for(Header header:headers){
//            String values = header.getValue();
//            String token = values.substring(0,values.indexOf(";",1));
//            String name =token.substring(0,token.indexOf("=")) ;
//            String value = token.substring(token.indexOf("=")+1,token.length());
//            BasicClientCookie basicClientCookie = new BasicClientCookie(name,value);
//            basicClientCookie.setPath("/");
//            if (name.equals("JSESSIONID")){
//                basicClientCookie.setDomain(".crm.xiaoshouyi.com");
//            }
//            basicClientCookie.setDomain(".xiaoshouyi.com");
//            cookieStore.addCookie(basicClientCookie);
//        }
//        return cookieStore;
//    }

    //读取excel将元素放入list
    public void getExcelParams(HttpClientTest httpClientTest) throws IOException, ParseException {
        FileInputStream fis = new FileInputStream(new File(PATH));
        //打开需要读取的文件
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //按照SHEET的名称读取一个电子表格
        XSSFSheet sheet = workbook.getSheet("sheet");
        //获取一个行的迭代器
        Iterator<Row> rowIterator = sheet.rowIterator();
        List pairList = new LinkedList();
        while(rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String parseExcel = this.parseExcel(cell);
                //放入list中
                pairList.add(parseExcel);
            }
            //发送请求
            httpClientTest.post(pairList,post);
            pairList.clear();
        }
        fis.close();
    }
    private String parseExcel(Cell cell) throws ParseException {
        String result;
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
                            .getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result=String.valueOf(sdf.parse(sdf.format(date)).getTime());
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case XSSFCell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                result = "";
            default:
                result = "";
                break;
        }
        return result;
    }
    /**
     * post方式提交表单（模拟用户登录请求）
     */
    public  void login(String url) {
        List formparams = new LinkedList();
        // 创建默认的httpClient实例.
        CloseableHttpClient httpclient = HttpClients.createDefault();
        // 创建httpPost
        HttpPost httppost = new HttpPost(url);
        // 创建参数队列
        formparams.add(new BasicNameValuePair("loginName", "liliang@hcis.com.cn"));
        formparams.add(new BasicNameValuePair("password", "cZ6568"));
        UrlEncodedFormEntity uefEntity;
        try {
            uefEntity = new UrlEncodedFormEntity(formparams, "UTF-8");
            httppost.setEntity(uefEntity);
            System.out.println("executing request " + httppost.getURI());
            CloseableHttpResponse response = httpclient.execute(httppost);
            try {
                HttpEntity entity = response.getEntity();
                headers=response.getHeaders("Set-Cookie");
                if (entity != null) {
                    System.out.println("--------------------------------------");
                    System.out.println("Response content: " + EntityUtils.toString(entity, "UTF-8"));
                    System.out.println("--------------------------------------");
                }
            } finally {
                response.close();
            }
        } catch (ClientProtocolException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭连接,释放资源
            try {
                httpclient.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    //post请求
    public void post(List formparams,String url) {
        List forms = new LinkedList();
        // 创建默认的httpClient实例.
        CloseableHttpClient httpClient = HttpClients.createDefault();
        // 创建httpPost
        HttpPost httppost = new HttpPost(url);
        // 创建参数队列
        if (formparams.size()==8){
            forms.add(new BasicNameValuePair("paramMap['money']", (String) formparams.get(0)));
            forms.add(new BasicNameValuePair("paramMap['expenseType']",(String) formparams.get(1)));
            forms.add(new BasicNameValuePair("paramMap['occurrenceDate']", (String) formparams.get(2)));
            forms.add(new BasicNameValuePair("paramMap['relateEntityId']", (String) formparams.get(3)));
            forms.add(new BasicNameValuePair("paramMap['relateEntity']", (String) formparams.get(4)));
            forms.add(new BasicNameValuePair("paramMap['dimDepart']", (String) formparams.get(5)));
            forms.add(new BasicNameValuePair("paramMap['dbcVarchar1']", (String) formparams.get(6)));
            forms.add(new BasicNameValuePair("paramMap['dbcVarchar2']", (String) formparams.get(7)));
        }
        UrlEncodedFormEntity uefEntity;
        try {
            uefEntity = new UrlEncodedFormEntity(forms, "UTF-8");
            httppost.setEntity(uefEntity);
            //组装cookie
            String result="";
            for (Header header: headers){
               result = result+header.getValue()+";";
            }
            httppost.addHeader("Cookie",result);
            System.out.println("executing request " + httppost.getURI());
            CloseableHttpResponse response = httpClient.execute(httppost);
            try {
                HttpEntity entity = response.getEntity();
                if (entity != null) {
                    System.out.println("--------------------------------------");
                    System.out.println("Response content: " + EntityUtils.toString(entity, "UTF-8"));
                    System.out.println("--------------------------------------");
                }
            } finally {
                response.close();
            }
        } catch (ClientProtocolException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭连接,释放资源
            try {
                httpClient.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    public static void main(String[] args) throws IOException, ParseException {
        HttpClientTest httpClientTest = new HttpClientTest();
        httpClientTest.login(urlLogin);
        httpClientTest.getExcelParams(httpClientTest);
    }

}
