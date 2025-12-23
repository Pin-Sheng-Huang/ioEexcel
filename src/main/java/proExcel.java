import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @description:
 * @Creator: 阿昇
 * @CreateTime: 2025-12-23 13:14
 */

public class proExcel {
    public static Logger logger = LogManager.getLogger(proExcel.class);
    public static void main(String[] args) {
            /*
        CREATE TABLE `t_import_member_temp`  (
  `USER_ACCOUNT` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL COMMENT '用户账号',
  `USER_PASSWORD` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '用户密码',
  `USER_NAME` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '姓名',
  `REGISTER_SOURCE` tinyint(1) NULL DEFAULT NULL COMMENT '注册来源 1:PC 2:IOS 3:安卓 4:H5',
  `REGISTER_SOURCE_NAME` varchar(30) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '设备ID',
  `BUNDLE_VERSION_ID` varchar(30) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '代理商',
  `AMOUNT` decimal(20, 2) NULL DEFAULT NULL COMMENT '初始金额',
  `DATASOURCE_KEY` varchar(30) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '分库键',
  `CREATION_TIME` datetime NULL DEFAULT NULL COMMENT '创建时间',
  `BATCH_NO` int NULL DEFAULT NULL COMMENT '批次号',
  `THREAD_NAME` varchar(40) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '线程名',
  `AMOUNT_STATUS` tinyint(1) NULL DEFAULT 0 COMMENT '金额状态，1为成功，0为失败',
  `EXCEPTION_INFO` varchar(2000) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL COMMENT '异常信息',
  PRIMARY KEY (`USER_ACCOUNT`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Dynamic;
        * * */
        //INSERT INTO `t_import_member_temp` VALUES ('07240as', NULL, '刘毅伟', 5, 'com.bty24.;iveios', 'tt9134', 0.00, NULL, NULL, NULL, NULL, 0, NULL);
        //INSERT INTO `8bet`.`t_import_member_temp`(`USER_ACCOUNT`, `USER_PASSWORD`, `USER_NAME`, `REGISTER_SOURCE`, `REGISTER_SOURCE_NAME`, `BUNDLE_VERSION_ID`, `AMOUNT`, `DAMA_VALUE`, `DATASOURCE_KEY`, `CREATION_TIME`, `BATCH_NO`, `THREAD_NAME`, `AMOUNT_STATUS`, `EXCEPTION_INFO`) 
        String sqlTemplate = "INSERT INTO `t_import_member_temp` VALUES (%s);";
        boolean skipFirstRow = true;
        int printRecordCount = 0; // 打印记录计数器
        List<String> sqlStatements = new ArrayList<>(); // 存储生成的 SQL 语句
        /*從excel寫入*/
        try (FileInputStream fileInputStream = new FileInputStream("C:\\Users\\USER\\Desktop\\导数据群的xlsx\\11.xlsx");
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            //test
           Sheet sheet = workbook.getSheetAt(1); // 获取 哪個工作表

            for (Row row : sheet) {
                // 如果需要跳过第一行，则将标志设置为 false，并继续下一行
                if (skipFirstRow) {
                    skipFirstRow = false;
                    continue;
                }

                StringBuilder values = new StringBuilder();
                boolean rowEmpty = true; // 行是否为空

                for (Cell cell : row) {
                    String cellValue = "";
                    DataFormatter formatter = new DataFormatter(); // 建議在外層建立一次重複使用
                    // 根据单元格类型读取数据并构建插入语句中的值
                    switch (cell.getCellType() ) {

                        case STRING:
                            String stringValue = cell.getStringCellValue();
                            if (stringValue != null && !stringValue.equalsIgnoreCase("NULL")) {
                                cellValue = "'" + stringValue + "'";
                            } else {
                                cellValue = stringValue;
                            }
                            break;
                        case NUMERIC: if (DateUtil.isCellDateFormatted(cell)) {
                             String dateStr = formatter.formatCellValue(cell); cellValue = "'" + dateStr + "'"; }
                        else {  String numStr = formatter.formatCellValue(cell);
                            if (numStr == null || numStr.trim().isEmpty()) { cellValue = "NULL"; }
                            else {
                                cellValue =  numStr ;
                            }
                            }
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        default:
                            // 处理其他类型
                    }

                    values.append(cellValue).append(", ");

                    // 如果单元格有数据，将行标记为非空
                    if (!cellValue.isEmpty()) {
                        rowEmpty = false;
                    }
                }

                // 如果整行数据为空，跳过该行
                if (rowEmpty) {
                    continue;
                }

                // 移除最后的逗号和空格
                if (values.length() > 0) {
                    values.setLength(values.length() - 2);
                }

                // 构建完整的插入语句
                String sql = String.format(sqlTemplate, values.toString());

                // 存储生成的 SQL 语句
                sqlStatements.add(sql);

                // 递增打印记录计数器
                printRecordCount++;
            }

            // 打印打印记录总数
            System.out.println("打印记录总数: " + printRecordCount);

            // 将 SQL 语句写入文件
            writeSqlToFile(sqlStatements, "C:\\Users\\USER\\Desktop\\导数据群的xlsx\\导出的sql文件.sql");



        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeSqlToFile(List<String> sqlStatements, String filePath) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {//高效写入文本的
            for (String sql : sqlStatements) {
                writer.write(sql);
                writer.newLine();
            }
            System.out.println("SQL 语句已保存到文件: " + filePath);

            // 在本地端打开文件
            try {
                // 使用默认关联的应用程序打开文件
                ProcessBuilder processBuilder = new ProcessBuilder("cmd.exe", "/c", "start", filePath);
                processBuilder.start();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
 }

