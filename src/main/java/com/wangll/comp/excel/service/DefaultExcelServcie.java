package com.wangll.comp.excel.service;

import com.beyond.framework.utils.DateUtils;
import com.wangll.comp.excel.bean.ExcelColBean;
import com.wangll.comp.excel.utils.ExcelUtils;
import com.wangll.comp.excel.utils.anno.ExcelSupport;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.view.document.AbstractExcelView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @package: com.beyond.transafemrg.resperson.service.
 * Created by ll_wang on 16/4/1.
 */

@Service
public class DefaultExcelServcie<T> extends AbstractExcelView {

    @Autowired
    CodeLibraryService codeLibraryService;

    @Override
    protected void buildExcelDocument(Map<String, Object> map, HSSFWorkbook hssfWorkbook, HttpServletRequest httpServletRequest, HttpServletResponse httpServletResponse) throws Exception {
        List<T> dataList = (List<T>)map.get("data");
        String exprotName = (String)map.get("exprotName");
        String sheetName = (String)map.get("sheetName");
        //编码
        Map<String,List<Map<String,Object>>> codes = (Map<String,List<Map<String,Object>>>)map.get("codes");
        if (codes == null) {
            codes = codeLibraryService.getInitCodeLibary();
        }

        ExcelSupport anno = null;
        Class clazz = null;
        Method method = null;
        Method[] methods = null;
        Type returnType = null;
        T obj = null;
        Object value = null;
        HSSFCell cell = null;
        List<ExcelColBean> colList = null;
        HSSFCellStyle cellStyle = null;

        //创建工作薄
        HSSFSheet sheet = hssfWorkbook.createSheet(sheetName != null ? sheetName : "默认工作薄");

        if (dataList != null && dataList.size() > 0) {
            colList = ExcelUtils.resolveAnno(dataList.get(0));
        }

        if (colList != null) {
            for (int i = 0; i < colList.size(); i ++) {
                cell=getCell(sheet, 0, i);
                cell.setCellValue(colList.get(i).getTitle());
            }
            clazz = dataList.get(0).getClass();
            methods = clazz.getMethods();
        }

        for (int i = 0 ; i < dataList.size(); i ++) {
            obj = dataList.get(i);
            if (colList != null) {
                for (int j = 0; j < colList.size(); j ++) {
                    String fieldName = colList.get(j).getName();
                    String methodStr = ExcelUtils.strToMethod(fieldName, "get");
                    cell = getCell(sheet, i + 1, j);
                    //匹配getXxxx方法
                    if (ExcelUtils.clazzHasMethod(methods, methodStr)) {
                        method = clazz.getMethod(methodStr);
                    //匹配getxXXXX方法
                    } else if(ExcelUtils.clazzHasMethod(methods, "get" + fieldName)) {
                        method = clazz.getMethod("get" + fieldName);
                    //匹配字段名相同的方法
                    } else if(ExcelUtils.clazzHasMethod(methods, fieldName)) {
                        method = clazz.getMethod(fieldName);
                    }

                    returnType = method.getReturnType();
                    value = method.invoke(obj, null);
                    if (value == null || "null".equals(value)) {
                        value = "";
                    }
                    if (returnType == Date.class && value instanceof Date && !"".equals(value)) {
                        value = DateUtils.parseDateToStr((Date)value, colList.get(j).getFormat());
                    }
                    //从编码表查对应的值
                    value = ExcelUtils.getBmCode(codes, fieldName, String.valueOf(value));
                    cell.setCellValue( String.valueOf(value) );

                }
            }
        }

        String filename = exprotName != null ? exprotName : "表格.xls";//设置下载时客户端Excel的名称
        filename = ExcelUtils.encodeFilename(filename, httpServletRequest);//处理中文文件名
        httpServletResponse.setContentType("application/vnd.ms-excel");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + filename);
        OutputStream ouputStream = httpServletResponse.getOutputStream();
        hssfWorkbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();
    }
}
