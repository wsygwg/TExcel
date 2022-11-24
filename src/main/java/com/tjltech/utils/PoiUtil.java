package com.tjltech.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

/**
 * POI工具类
 *
 * @author zhongfeng.fang
 * @date 2020/04/16
 */
public class PoiUtil {

    // Excel 2003 版本
    public static final String EXCEL_XLS = "xls";
    // Excel 2007 以上版本
    public static final String EXCEL_XLSX = "xlsx";

    /**
     * 判断是否是excel2003文件
     *
     * @param filePath filePath
     * @return 是否是excel2003文件
     */
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * 判断是否是excel2007文件
     *
     * @param filePath filePath
     * @return 是否是excel2007文件
     */
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 判断是否是excel文件
     *
     * @param filePath filePath
     * @return 是否是excel文件
     */
    public static boolean validateExcel(String filePath) {
        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {
            return false;
        }
        return true;
    }


    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param fileType fileType
     * @param in       in
     * @return Workbook
     */
    public static Workbook getWorkbook(String fileType, InputStream in) {
        Workbook workbook = null;
        try {
            if (fileType.equals(EXCEL_XLS)) {
                workbook = new HSSFWorkbook(in);
            } else if (fileType.equals(EXCEL_XLSX)) {
                workbook = new XSSFWorkbook(in);
            }
            in.close();
            if (null == workbook) {
                workbook = new XSSFWorkbook(in);
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        return workbook;
    }

    /**
     * 获取excel的第一个sheet的总行数，并且根据有数据的第一行获得总列数
     *
     * @param workBook workBook
     * @return excel的行数和列数
     */
    public static Map<String, Integer> getRowAndColumnCount(Workbook workBook) {

        Map<String, Integer> returnValue = new HashMap<String, Integer>();
        // 得到第一个shell
        Sheet sheet = workBook.getSheetAt(0);
        // 得到Excel的行数
        Integer totalRows = sheet.getPhysicalNumberOfRows();

        returnValue.put("totalRows", totalRows);

        Integer totalColumns = 0;
        // 得到Excel的列数(前提是有行数)
        if (totalRows > 1 && sheet.getRow(0) != null) {
            totalColumns = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        returnValue.put("totalColumns", totalColumns);

        return returnValue;
    }

    /**
     * 处理导入小数点
     *
     * @param cell cell
     * @return 去掉小数点后的0
     */
    public static Integer numOfImport(Cell cell) {
        DecimalFormat df = new DecimalFormat("0");
        String value = df.format(cell.getNumericCellValue());

        String[] str = value.split("\\.");
        if (str.length > 1) {
            String str1 = str[1];
            int m = Integer.parseInt(str1);
            if (m == 0) {
                return m;
            } else {
                return null;
            }
        } else {
            return Integer.parseInt(value);
        }
    }

    /**
     * 下载excel，设定OutputStream
     *
     * @param excelOutputPath excelOutputPath
     * @param outputStream    outputStream
     * @param workbook        workbook
     * @param fileName        fileName
     * @param fileType        fileType
     * @return filePath
     */
    public static String setOutputStream(String excelOutputPath, OutputStream outputStream, Workbook workbook,
                                         String fileName, String fileType) {
        // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
        String filePath = "";
        try {
            filePath = excelOutputPath + fileName + System.currentTimeMillis() + fileType;
            outputStream = new FileOutputStream(excelOutputPath + fileName + System.currentTimeMillis() + fileType);
            workbook.write(outputStream);
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return filePath;
    }

    /**
     * 给sheet页，添加下拉列表
     *
     * @param workbook    excel文件，用于添加Name
     * @param targetSheet 级联列表所在sheet页
     * @param options     级联数据 ['百度','阿里巴巴']
     * @param column      下拉列表所在列
     * @param fromRow     下拉限制开始行
     * @param endRow      下拉限制结束行
     */
    public static void addPullDownListToSheet(Workbook workbook, Sheet targetSheet, String[] options, int column,
                                              int fromRow, int endRow) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) targetSheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
                .createExplicitListConstraint(options);
        CellRangeAddressList addressList = new CellRangeAddressList(fromRow, endRow, column, column);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
        validation.setShowErrorBox(true);
        validation.setSuppressDropDownArrow(true);
        validation.setEmptyCellAllowed(false);
        validation.setShowPromptBox(true);
        validation.createPromptBox("提示", "只能选择下拉框里面的数据");
        targetSheet.addValidationData(validation);
    }

    /**
     * 下载EXCEL
     *
     * @param request      request
     * @param response     response
     * @param filePath     filePath
     * @param fileName     fileName
     * @param outputStream outputStream
     */
//    public static void downLoadExcel(HttpServletRequest request, HttpServletResponse response, String filePath,
//                                     String fileName, OutputStream outputStream) {
//
//        BufferedInputStream bis = null;
//        OutputStream os = null;
//        try {
//            response.reset();
//            String userAgent = request.getHeader("User-Agent");
//            if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
//                fileName = URLEncoder.encode(fileName, "utf-8");
//            } else {
//                fileName = new String(fileName.getBytes("utf-8"), "ISO-8859-1");
//            }
//            response.setHeader("Content-disposition", "attachment; filename=\"" + fileName + "\"");
//            response.setContentType("application/x-msdownload;charset=utf-8");
//            response.setCharacterEncoding("utf-8");
//
//            byte[] buff = new byte[MagicNumberConstants.ONE_ZERO_TWO_FOUR];
//            os = response.getOutputStream();
//            bis = new BufferedInputStream(new FileInputStream(filePath));
//            int i = 0;
//            while ((i = bis.read(buff)) != MagicNumberConstants.MINUS_ONE) {
//                os.write(buff, 0, i);
//                os.flush();
//            }
//            File file = new File(filePath);
//            if (file.exists() && file.isFile()) {
//                file.delete();
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        } finally {
//            try {
//                if (outputStream != null) {
//                    outputStream.flush();
//                    outputStream.close();
//                }
//                if (bis != null) {
//                    bis.close();
//                }
//                if (os != null) {
//                    os.close();
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }

    /**
     * 设定单元格样式
     *
     * @param workbook         workbook
     * @param cellStyleContent cellStyleContent
     * @return CellStyle
     */
    public static CellStyle setCellDataStyle(Workbook workbook, String cellStyleContent) {
        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyle cellStyle = workbook.createCellStyle();
        short style = createHelper.createDataFormat().getFormat(cellStyleContent);
        cellStyle.setDataFormat(style);
        return cellStyle;
    }

    /**
     * 复制单元格
     *
     * @param srcCell       控制是否复制单元格的内容
     * @param desCell       控制是否复制单元格的内容
     * @param copyValueFlag 控制是否复制单元格的内容
     * @param copyStyleFlag 控制是否复制样式
     */
//    public static void copyCell(Cell srcCell, Cell desCell, boolean copyValueFlag, boolean copyStyleFlag) {
//        Workbook srcBook = srcCell.getSheet().getWorkbook();
//        Workbook desBook = desCell.getSheet().getWorkbook();
//
//        // 复制样式
//        // 如果是同一个excel文件内，连带样式一起复制
//        if (srcBook == desBook && copyStyleFlag) {
//            // 同文件，复制引用
//            desCell.setCellStyle(srcCell.getCellStyle());
//        }
//
//        // 复制评论
//        if (srcCell.getCellComment() != null) {
//            desCell.setCellComment(srcCell.getCellComment());
//        }
//
//        // 复制内容
//        desCell.setCellType(srcCell.getCellTypeEnum());
//
//        if (copyValueFlag) {
//            switch (srcCell.getCellTypeEnum()) {
//                case STRING:
//                    desCell.setCellValue(srcCell.getStringCellValue());
//                    break;
//                case NUMERIC:
//                    desCell.setCellValue(srcCell.getNumericCellValue());
//                    break;
//                case FORMULA:
//                    desCell.setCellFormula(srcCell.getCellFormula());
//                    break;
//                case BOOLEAN:
//                    desCell.setCellValue(srcCell.getBooleanCellValue());
//                    break;
//                case ERROR:
//                    desCell.setCellValue(srcCell.getErrorCellValue());
//                    break;
//                case BLANK:
//                    // nothing to do
//                    break;
//                default:
//                    break;
//            }
//        }
//    }

    /**
     * 判断单元格的值是否是日期型
     *
     * @param cell cell
     * @return boolean
     */
    public static boolean isCellDateFormatted(Cell cell) {
        if (cell == null) {
            return false;
        }
        boolean bDate = false;
        double d;
        try {
            d = cell.getNumericCellValue();
        } catch (Exception ex) {
            return false;
        }
        if (isValidExcelDate(d)) {
            CellStyle style = cell.getCellStyle();
            if (style == null) {
                return false;
            }
            int i = style.getDataFormat();
            String f = style.getDataFormatString();
            bDate = isADateFormat(i, f);
        }
        return bDate;
    }

    /**
     * 判断EXCEL单元格是否是日期类型
     *
     * @param formatIndex  formatIndex
     * @param formatString formatString
     * @return boolean
     */
    public static boolean isADateFormat(int formatIndex, String formatString) {
        if ((formatString == null) || (formatString.length() == 0)) {
            return false;
        }
        String fs = formatString;
        // 下面这一行是自己手动添加的 以支持汉字格式
        fs = fs.replaceAll("[\"|\']", "").replaceAll("[年|月|日|时|分|秒|毫秒|微秒]", "");
        fs = fs.replaceAll("\\\\-", "-");
        fs = fs.replaceAll("\\\\,", ",");
        fs = fs.replaceAll("\\\\.", ".");
        fs = fs.replaceAll("\\\\ ", " ");
        fs = fs.replaceAll(";@", "");
        fs = fs.replaceAll("^\\[\\$\\-.*?\\]", "");
        fs = fs.replaceAll("^\\[[a-zA-Z]+\\]", "");
        return (fs.matches("^[yYmMdDhHsS\\-/,. :]+[ampAMP/]*$"));
    }

    /**
     * 判断日期值是否合理
     *
     * @param value value
     * @return boolean
     */
    public static boolean isValidExcelDate(double value) {
        return (value > MagicNumberConstants.EXCEL_DATE_VALUE_RANGE);
    }


//    public static boolean isEmptyRow(Row row) {
//        if (row == null || row.toString().isEmpty()) {
//            return true;
//        } else {
//            Iterator<Cell> it = row.iterator();
//            boolean isEmpty = true;
//            while (it.hasNext()) {
//                Cell cell = it.next();
//                if (cell.getCellType() != CellType.BLANK) {
//                    isEmpty = false;
//                    break;
//                }
//            }
//            return isEmpty;
//        }
//    }

}
