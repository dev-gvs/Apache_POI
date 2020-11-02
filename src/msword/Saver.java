package msword;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Saver {

    private static final String dir
            = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
            + System.getProperty("file.separator");
    private static final Logger logger = Logger.getLogger(Saver.class);

    public static void saveDoc(String number, String fio, String address) {
        new Thread(() -> {
            // Чтение из шаблона в переменную doc
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "receipt_template.doc")) {
                doc = new HWPFDocument(fis);
                logger.debug("Read .doc template");
            } catch (Exception e) {
                logger.error(e);
            }

            // Замена в переменной doc данных
            try {
                doc.getRange().replaceText("$НОМЕРполучателя", number);
                logger.debug("Replaced '$НОМЕРполучателя' with '" + number + "'");
                doc.getRange().replaceText("$ФИОплательщика", fio);
                logger.debug("Replaced '$ФИОплательщика' with '" + fio + "'");
                doc.getRange().replaceText("$АДРЕСплательщика", address);
                logger.debug("Replaced '$АДРЕСплательщика' with '" + address + "'");
            } catch (Exception e) {
                logger.error(e);
            }

            // Сохранение переменной doc в новый файл
            try (FileOutputStream fos = new FileOutputStream(dir + "receipt.doc")) {
                doc.write(fos);
                logger.debug("Wrote result to receipt.doc");
                // Открытие файла внешней программой
                Desktop.getDesktop().open(new File(dir + "receipt.doc"));
                logger.debug("Opened receipt.doc");
            } catch (Exception e) {
                logger.error(e);
            }
        }).start();
    }

    public static void saveXls(String fio, String name, String address, String sum, String sumUsl) {
        // Чтение из шаблона в переменную xls
        HSSFWorkbook xls = null;

        try (FileInputStream fis = new FileInputStream(dir + "receipt_template.xls")) {
            xls = new HSSFWorkbook(fis);
            logger.debug("Read .xls template");
        } catch (Exception e) {
            logger.error(e);
        }

        // Первый лист документа
        HSSFSheet sheet = null;
        sheet = xls.getSheetAt(0);

        // Замена в переменной xls данных
        try {
            replaceCellText(sheet, 1, 2, fio);
            logger.debug("Added '" + fio + "' to sheet at (1, 2)");
            replaceCellText(sheet, 12, 3, name);
            logger.debug("Added '" + name + "' to sheet at (12, 3)");
            replaceCellText(sheet, 13, 3, address);
            logger.debug("Added '" + address + "' to sheet at (13, 3)");
            replaceCellText(sheet, 14, 3, sum);
            logger.debug("Added '" + sum + "' to sheet at (14, 3)");
            replaceCellText(sheet, 14, 8, sumUsl);
            logger.debug("Added '" + sumUsl + "' to sheet at (14, 8)");
            int totalSum = Integer.parseInt(sum) + Integer.parseInt(sumUsl);
            replaceCellText(sheet, 15, 3, Integer.toString(totalSum));
            logger.debug("Added '" + totalSum + "' to sheet at (15, 3)");
        } catch (Exception e) {
            logger.error(e);
        }

        // Сохранение переменной doc в новый файл
        try (FileOutputStream fos = new FileOutputStream(dir + "receipt.xls")) {
            xls.write(fos);
            logger.debug("Wrote result to receipt.xls");
            // Открытие файла внешней программой
            Desktop.getDesktop().open(new File(dir + "receipt.xls"));
            logger.debug("Opened receipt.xls");
        } catch (Exception e) {
            logger.error(e);
        }
    }

    public static void saveDocx(String number, String fio, String address) {
        new Thread(() -> {
            // Чтение из шаблона в переменную doc
            XWPFDocument docx = null;
            try (FileInputStream fis = new FileInputStream(dir + "receipt_template.docx")) {
                docx = new XWPFDocument(OPCPackage.open(fis));
                logger.debug("Read .docx template");
            } catch (Exception e) {
                logger.error(e);
            }

            // Замена в переменной doc данных
            try {
                replaceText(docx, "НОМЕРполучателя", number);
                logger.debug("Replaced 'НОМЕРполучателя' with '" + number + "'");
                replaceText(docx, "ФИОплательщика", fio);
                logger.debug("Replaced 'ФИОплательщика' with '" + fio + "'");
                replaceText(docx, "АДРЕСплательщика", address);
                logger.debug("Replaced 'АДРЕСплательщика' with '" + address + "'");
            } catch (Exception e) {
                logger.error(e);
            }

            // Сохранение переменной doc в новый файл
            try (FileOutputStream fos = new FileOutputStream(dir + "receipt.docx")) {
                docx.write(fos);
                logger.debug("Wrote result to receipt.docx");
                // Открытие файла внешней программой
                Desktop.getDesktop().open(new File(dir + "receipt.docx"));
                logger.debug("Opened receipt.docx");
            } catch (Exception e) {
                logger.error(e);
            }
        }).start();
    }

    public static void saveXlsx(String fio, String name, String address, String sum, String sumUsl) {
        // Чтение из шаблона в переменную xlsx
        XSSFWorkbook xlsx = null;

        try (FileInputStream fis = new FileInputStream(dir + "receipt_template.xlsx")) {
            xlsx = new XSSFWorkbook(OPCPackage.open(fis));
            logger.debug("Read .xlsx template");
        } catch (Exception e) {
            logger.error(e);
        }

        // Первый лист документа
        XSSFSheet sheet = null;
        sheet = xlsx.getSheetAt(0);

        // Замена в переменной xls данных
        try {
            replaceCellText(sheet, 1, 2, fio);
            logger.debug("Added '" + fio + "' to sheet at (1, 2)");
            replaceCellText(sheet, 12, 3, name);
            logger.debug("Added '" + name + "' to sheet at (12, 3)");
            replaceCellText(sheet, 13, 3, address);
            logger.debug("Added '" + address + "' to sheet at (13, 3)");
            replaceCellText(sheet, 14, 3, sum);
            logger.debug("Added '" + sum + "' to sheet at (14, 3)");
            replaceCellText(sheet, 14, 8, sumUsl);
            logger.debug("Added '" + sumUsl + "' to sheet at (14, 8)");
            int totalSum = Integer.parseInt(sum) + Integer.parseInt(sumUsl);
            replaceCellText(sheet, 15, 3, Integer.toString(totalSum));
            logger.debug("Added '" + totalSum + "' to sheet at (15, 3)");
        } catch (Exception e) {
            logger.error(e);
        }

        // Сохранение переменной doc в новый файл
        try (FileOutputStream fos = new FileOutputStream(dir + "receipt.xlsx")) {
            xlsx.write(fos);
            logger.debug("Wrote result to receipt.xlsx");
            // Открытие файла внешней программой
            Desktop.getDesktop().open(new File(dir + "receipt.xlsx"));
            logger.debug("Opened receipt.xlsx");
        } catch (Exception e) {
            logger.error(e);
        }
    }

    private static void replaceCellText(XSSFSheet sheet, int row, int cell, String text) {
        sheet.getRow(row).getCell(cell).setCellValue(text);
    }

    private static void replaceCellText(HSSFSheet sheet, int row, int cell, String text) {
        sheet.getRow(row).getCell(cell).setCellValue(text);
    }

    private static void replaceText(XWPFDocument docx, String textToFind, String textToReplace) {
        for (XWPFTable tbl : docx.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains(textToFind)) {
                                text = text.replace(textToFind, textToReplace);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }
}
