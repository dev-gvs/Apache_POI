package msword;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;

public class ReceiptWord extends javax.swing.JFrame {
    
    private static final Logger logger = Logger.getLogger(ReceiptWord.class);

    String dir;
    Runnable saveDocTask;
    Runnable saveXlsTask;

    /**
     * Creates new form ReceiptWord
     */
    public ReceiptWord() {
        initComponents();
        logger.debug("Initialized components");

        this.dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                + System.getProperty("file.separator");

        this.saveDocTask = () -> {
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
                doc.getRange().replaceText("$НОМЕРполучателя", jTextFieldNumber.getText());
                logger.debug("Replaced '$НОМЕРполучателя' with '" + jTextFieldNumber.getText() + "'");
                doc.getRange().replaceText("$ФИОплательщика", jTextFieldName.getText());
                logger.debug("Replaced '$ФИОплательщика' with '" + jTextFieldName.getText() + "'");
                doc.getRange().replaceText("$АДРЕСплательщика", jTextFieldAddress.getText());
                logger.debug("Replaced '$АДРЕСплательщика' with '" + jTextFieldAddress.getText() + "'");
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
        };

        this.saveXlsTask = () -> {
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

            // Замена в переменной doc данных
            try {
                sheet.getRow(1).getCell(2).setCellValue(jTextFieldXlsReceiverFIO.getText());
                logger.debug("Added '" + jTextFieldXlsReceiverFIO.getText() + "' to sheet at (1, 2)");
                sheet.getRow(12).getCell(3).setCellValue(jTextFieldXlsName.getText());
                logger.debug("Added '" + jTextFieldXlsName.getText() + "' to sheet at (12, 3)");
                sheet.getRow(13).getCell(3).setCellValue(jTextFieldXlsAddress.getText());
                logger.debug("Added '" + jTextFieldXlsAddress.getText() + "' to sheet at (13, 3)");
                sheet.getRow(14).getCell(3).setCellValue(jTextFieldXlsSum.getText());
                logger.debug("Added '" + jTextFieldXlsSum.getText() + "' to sheet at (14, 3)");
                sheet.getRow(14).getCell(8).setCellValue(jTextFieldXlsSumUsl.getText());
                logger.debug("Added '" + jTextFieldXlsSumUsl.getText() + "' to sheet at (14, 8)");
                int sum = Integer.parseInt(jTextFieldXlsSum.getText()) +
                        Integer.parseInt(jTextFieldXlsSumUsl.getText());
                sheet.getRow(15).getCell(3).setCellValue(sum);
                logger.debug("Added '" + sum + "' to sheet at (15, 3)");
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
        };
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPanel1 = new javax.swing.JPanel();
        jTextFieldName = new javax.swing.JTextField();
        jTextFieldAddress = new javax.swing.JTextField();
        jButtonSave = new javax.swing.JButton();
        jTextFieldNumber = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        jTextFieldXlsReceiverFIO = new javax.swing.JTextField();
        jTextFieldXlsName = new javax.swing.JTextField();
        jTextFieldXlsAddress = new javax.swing.JTextField();
        jButtonXlsSave = new javax.swing.JButton();
        jTextFieldXlsSum = new javax.swing.JTextField();
        jTextFieldXlsSumUsl = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Квитанции");
        setResizable(false);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jTabbedPane1.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N

        jPanel1.setLayout(null);

        jTextFieldName.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        jPanel1.add(jTextFieldName);
        jTextFieldName.setBounds(500, 270, 530, 25);

        jTextFieldAddress.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        jTextFieldAddress.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextFieldAddressActionPerformed(evt);
            }
        });
        jPanel1.add(jTextFieldAddress);
        jTextFieldAddress.setBounds(500, 300, 530, 25);

        jButtonSave.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSave.setText("Экспортировать в .doc");
        jButtonSave.setToolTipText("");
        jButtonSave.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveActionPerformed(evt);
            }
        });
        jPanel1.add(jButtonSave);
        jButtonSave.setBounds(20, 260, 260, 80);

        jTextFieldNumber.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        jPanel1.add(jTextFieldNumber);
        jTextFieldNumber.setBounds(630, 190, 400, 25);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/msword/receipt.png"))); // NOI18N
        jPanel1.add(jLabel1);
        jLabel1.setBounds(4, 4, 1040, 430);

        jTabbedPane1.addTab(".doc", jPanel1);

        jPanel2.setLayout(null);

        jTextFieldXlsReceiverFIO.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        jPanel2.add(jTextFieldXlsReceiverFIO);
        jTextFieldXlsReceiverFIO.setBounds(295, 33, 680, 25);
        jPanel2.add(jTextFieldXlsName);
        jTextFieldXlsName.setBounds(420, 250, 560, 25);

        jTextFieldXlsAddress.setToolTipText("");
        jPanel2.add(jTextFieldXlsAddress);
        jTextFieldXlsAddress.setBounds(420, 280, 560, 25);

        jButtonXlsSave.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonXlsSave.setText("Экспортировать в .xls");
        jButtonXlsSave.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonXlsSaveActionPerformed(evt);
            }
        });
        jPanel2.add(jButtonXlsSave);
        jButtonXlsSave.setBounds(20, 260, 260, 80);

        jTextFieldXlsSum.setText("0");
        jTextFieldXlsSum.setToolTipText("");
        jPanel2.add(jTextFieldXlsSum);
        jTextFieldXlsSum.setBounds(420, 310, 60, 25);

        jTextFieldXlsSumUsl.setText("0");
        jTextFieldXlsSumUsl.setToolTipText("");
        jPanel2.add(jTextFieldXlsSumUsl);
        jTextFieldXlsSumUsl.setBounds(720, 310, 75, 25);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/msword/receipt_xls.png"))); // NOI18N
        jPanel2.add(jLabel2);
        jLabel2.setBounds(20, 0, 980, 430);

        jTabbedPane1.addTab(".xls", jPanel2);

        getContentPane().add(jTabbedPane1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 1040, 460));
        jTabbedPane1.getAccessibleContext().setAccessibleName(".doc");

        setBounds(0, 0, 1058, 504);
    }// </editor-fold>//GEN-END:initComponents

    private void jButtonSaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveActionPerformed
        logger.debug("Export to .doc button has been pressed");
        
        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        logger.debug("Starting saveDocTask");
        new Thread(saveDocTask).start();

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveActionPerformed

    private void jTextFieldAddressActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextFieldAddressActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextFieldAddressActionPerformed

    private void jButtonXlsSaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonXlsSaveActionPerformed
        logger.debug("Export to .xls button has been pressed");
        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        logger.debug("Starting saveXlsTask");
        new Thread(saveXlsTask).start();

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonXlsSaveActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ReceiptWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ReceiptWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ReceiptWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ReceiptWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptWord().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButtonSave;
    private javax.swing.JButton jButtonXlsSave;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTextField jTextFieldAddress;
    private javax.swing.JTextField jTextFieldName;
    private javax.swing.JTextField jTextFieldNumber;
    private javax.swing.JTextField jTextFieldXlsAddress;
    private javax.swing.JTextField jTextFieldXlsName;
    private javax.swing.JTextField jTextFieldXlsReceiverFIO;
    private javax.swing.JTextField jTextFieldXlsSum;
    private javax.swing.JTextField jTextFieldXlsSumUsl;
    // End of variables declaration//GEN-END:variables
}
