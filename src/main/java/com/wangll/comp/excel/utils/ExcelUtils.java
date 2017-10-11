package com.wangll.comp.excel.utils;

import com.beyond.appbase.excel.bean.ExcelColBean;
import com.beyond.appbase.excel.exception.ImportExcelException;
import com.beyond.appbase.excel.utils.anno.ExcelSupport;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.util.StringUtils;

import javax.mail.internet.MimeUtility;
import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.net.URLEncoder;
import java.text.ParseException;
import java.util.*;

/**
 * @Description: Excel导入导出辅助工具类
 * @package: com.beyond.transafemrg.common.utils.
 * Created by ll_wang on 16/4/1.
 */
public class ExcelUtils {
    /**
     * 设置下载文件中文件的名称
     *
     * @param filename
     * @param request
     * @return
     */
    public static String encodeFilename(String filename, HttpServletRequest request) {
        /**
         * 获取客户端浏览器和操作系统信息
         * 在IE浏览器中得到的是：User-Agent=Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Maxthon; Alexa Toolbar)
         * 在Firefox中得到的是：User-Agent=Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.7.10) Gecko/20050717 Firefox/1.0.6
         */
        String agent = request.getHeader("USER-AGENT");
        try {
            if ((agent != null) && (-1 != agent.indexOf("MSIE"))) {
                String newFileName = URLEncoder.encode(filename, "UTF-8");
                newFileName = StringUtils.replace(newFileName, "+", "%20");
                if (newFileName.length() > 150) {
                    newFileName = new String(filename.getBytes("GB2312"), "ISO8859-1");
                    newFileName = StringUtils.replace(newFileName, " ", "%20");
                }
                return newFileName;
            }
            if ((agent != null) && (-1 != agent.indexOf("Mozilla")))
                return MimeUtility.encodeText(filename, "UTF-8", "B");

            return filename;
        } catch (Exception ex) {
            return filename;
        }
    }

    /**
     * 解析注解
     * @param obj
     * @return
     */
    public static List<ExcelColBean> resolveAnno(Object obj) {
        ExcelSupport anno = null;
        List<ExcelColBean> beans = new ArrayList<ExcelColBean>();
        ExcelColBean col = null;
        Class clazz = obj.getClass();
        Field[] fields = clazz.getDeclaredFields();
        for(Field field : fields) {
            col = new ExcelColBean();
            anno = field.getDeclaredAnnotation(ExcelSupport.class);
            if (anno != null && anno.use()) {
                col.setName(field.getName());
                col.setTitle(anno.name());
                col.setSort(anno.sort());
                col.setWidth(anno.cellWidth());
                col.setWrap(anno.wrap());
                col.setCode(anno.code());
                col.setFormat(anno.format());
                beans.add(col);
            }
        }

        // 对列表排序
        Collections.sort(beans, new Comparator<ExcelColBean>() {
            public int compare(ExcelColBean arg0, ExcelColBean arg1) {
                int hits0 = arg0.getSort();
                int hits1 = arg1.getSort();
                if (hits1 > hits0) {
                    return -1;
                } else if (hits1 == hits0) {
                    return 0;
                } else {
                    return 1;
                }
            }
        });
        return beans;
    }

    /**
     *  字段名转换为方法名
     * @param field 字段
     * @param prefix 前缀
     * @param suffix 后缀
     * @return
     */
    public static String strToMethod(String field, String prefix, String suffix) {
        field = toUpperCaseFirstOne(field);
        return prefix + field + suffix;
    }
    public static String strToMethod(String field, String prefix) {
        return strToMethod(field, prefix, "");
    }

    //首字母转大写
    public static String toUpperCaseFirstOne(String s) {
        if(Character.isUpperCase(s.charAt(0)))
            return s;
        else
            return (new StringBuilder()).append(Character.toUpperCase(s.charAt(0))).append(s.substring(1)).toString();
    }

    //集合中是否有名为methodStr的方法名
    public static boolean clazzHasMethod(Method[] methods, String methodStr) {
        for(Method method : methods) {
            if (method.getName().equals(methodStr)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 从编码表中获取对应的值
     * @param codes 集合
     * @param field 字段对应的编码表
     * @param value 字符串对应的编码
     * @return
     */
    public static String getBmCode(Map<String,List<Map<String,Object>>> codes, String field, String value) {
        List<Map<String,Object>> bmCode = codes.get(field);
        if (field == null) {
            field = "";
        }
        if (bmCode != null) {
            for (Map<String,Object> code: bmCode) {
                if (value.equals(code.get("code"))) {
                    return (String) code.get("text");
                }
            }
        }
        return value;
    }

    /**
     * 从编码表中获取对应的值
     * @param codes 集合
     * @param field 字段对应的编码表
     * @param value 字符串对应的编码
     * @return
     */
    public static String getBmText(Map<String,List<Map<String,Object>>> codes, String field, String value) {
        List<Map<String,Object>> bmCode = codes.get(field);
        if (field == null) {
            field = "";
        }
        if (bmCode != null) {
            for (Map<String,Object> code: bmCode) {
                if (value.equals(code.get("text"))) {
                    return (String) code.get("code");
                }
            }
        }
        return value;
    }

    public static String dateFormat() {

        return null;
    }

    /**
     * 获取单元格对应的值
     * @param cell
     * @return
     */
    public static String getCellStringValue(HSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING://字符串类型
                cellValue = cell.getStringCellValue();
                if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
                    cellValue = "";
                break;
            case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_FORMULA: //公式
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                cellValue = " ";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }

    /**
     * 将字符串转换为制定的类型
     * @param value
     * @param type
     * @return
     */
    public static Object getValueType(String value, Type type) {
        if (type == String.class) {
            return value;
        } else if (type == Integer.class) {
            return Integer.parseInt(value);
        } else if (type == Short.class) {
            return Short.parseShort(value);
        } else if (type == Long.class) {
            return Long.parseLong(value);
        } else if (type == Double.class) {
            return Double.parseDouble(value);
        } else if (type == Float.class) {
            return Float.parseFloat(value);
        } else if (type == Date.class) {
            //TODO : 日期处理
            try {
                if (value != null && !"".equals(value)) {
                    return DateUtils.parseDate(value, new String[]{"yyyy-MM-dd HH:mm:ss"});
                }
            } catch (ParseException e) {
                e.printStackTrace();
            }
        } else if (type == Boolean.class) {
            return Boolean.parseBoolean(value);
        }
        return null;
    }

    /**
     * 将字符串转换为制定的类型
     * @param value
     * @param type
     * @return
     */
    public static Class<?> type2Clazz(Type type) {
        if (type == String.class) {
            return String.class;
        } else if (type == Integer.class) {
            return Integer.class;
        } else if (type == Short.class) {
            return Short.class;
        } else if (type == Long.class) {
            return Long.class;
        } else if (type == Double.class) {
            return Double.class;
        } else if (type == Float.class) {
            return Float.class;
        } else if (type == Date.class) {
            return Date.class;
        } else if (type == Boolean.class) {
            return Boolean.class;
        }
        return null;
    }

    /**
     * 将excel文件读取到list中
     * @param inputStream excel文件流
     * @param clazz 需操作的对象类型
     * @return
     * @throws IOException
     */
    public static List<Object> excelToList(InputStream inputStream, Class clazz, Map<String,List<Map<String,Object>>> codes) throws IOException, ImportExcelException {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
        List<Object> beans = new ArrayList<Object>();
        ExcelSupport anno = null;
        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        Method method;
        Object bean;
        HSSFRow hssfRow = null;
        HSSFCell brandIdHSSFCell = null;
        String value = "", fieldName = "";
        Type fieldType = null;
        List<Object> fieldArr, annoArr, useArr = new ArrayList<Object>();//创建字段数组和注解数组，以便最佳解析
        if(hssfWorkbook.getNumberOfSheets() == 0) {
            throw new ImportExcelException("没找到工作薄");
        }
        //循环工作薄
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            if(hssfSheet.getLastRowNum() == 0) {
                throw new ImportExcelException("工作薄：" + hssfSheet.getSheetName() + ",没有内容");
            }
            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                hssfRow = hssfSheet.getRow(rowNum);
                //获取第一行标题字段
                if (rowNum == 0) {
                    fieldArr = new ArrayList<Object>();
                    annoArr = new ArrayList<Object>();
                    for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                        value = getCellStringValue(hssfRow.getCell(i));
                        for (Field field : fields) {
                            anno = field.getDeclaredAnnotation(ExcelSupport.class);
                            fieldName = field.getName();
                            fieldType = field.getType();
                            if (fieldName.equals(value)) {
                                fieldArr.add(new Object[]{fieldName, fieldType});
                            }
                            if (anno != null && value.equals(anno.name())) {
                                annoArr.add(new Object[]{fieldName, fieldType, anno.code()});
                            }
                        }
                    }

                    //按匹配到字段较多的方式进行解析
                    if (fieldArr.size() > annoArr.size()) {
                        useArr = new ArrayList<Object>(Arrays.asList(new String[fieldArr.size()]));
                        Collections.copy(useArr, fieldArr);
                    } else {
                        useArr = new ArrayList<Object>(Arrays.asList(new String[annoArr.size()]));
                        Collections.copy(useArr, annoArr);
                    }
                    if(useArr.size() == 0) {
                        throw new ImportExcelException("该工作薄没有首行字段行");
                    }
                    continue;
                }
                try {
                    bean = clazz.newInstance();
                    //类反射实例对象
                    try {
                        for (int j = 0; j < useArr.size(); j ++) {
                            brandIdHSSFCell = hssfRow.getCell(j);
                            value = getCellStringValue(brandIdHSSFCell);
                            Object[] field = (Object[])useArr.get(j);
                            String methodStr = strToMethod((String)field[0], "set");
                            //反编码
                            if (field.length == 3 && (Boolean)field[2]) {
                                value = getBmText(codes, (String)field[0], value);
                            }
                            if(clazzHasMethod(methods, methodStr)) {
                                method = clazz.getDeclaredMethod(methodStr, type2Clazz((Type)field[1]));
                                method.invoke(bean, getValueType(value, (Type)field[1]));
                            } else if (clazzHasMethod(methods, (String)field[0])) {
                                method = clazz.getDeclaredMethod((String) field[0], type2Clazz((Type) field[1]));
                                method.invoke(bean, getValueType(value, (Class) field[1]));
                            } else if (clazzHasMethod(methods, "set" + field[0])) {
                                method = clazz.getDeclaredMethod("set" + field[0], type2Clazz((Type) field[1]));
                                method.invoke(bean, getValueType(value, (Class) field[1]));
                            }
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        throw new ImportExcelException(rowNum + "行," + "' "+ value +" '有误");
                    }
                    beans.add(bean);
                } catch (InstantiationException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }


            }
        }

        return beans;
    }

    public static void main(String[] args) {
        System.out.println(strToMethod("wangll", "get"));
    }


}
