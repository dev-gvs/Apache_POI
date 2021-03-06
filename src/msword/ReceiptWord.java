package msword;

import java.awt.Cursor;
import org.apache.log4j.Logger;

public class ReceiptWord extends javax.swing.JFrame {

    private static final Logger logger = Logger.getLogger(ReceiptWord.class);

    /**
     * Creates new form ReceiptWord
     */
    public ReceiptWord() {
        initComponents();
        logger.debug("Initialized components");
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
        jButtonSaveDoc = new javax.swing.JButton();
        jButtonSaveDocx = new javax.swing.JButton();
        jTextFieldNumber = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        jTextFieldXlsReceiverFIO = new javax.swing.JTextField();
        jTextFieldXlsName = new javax.swing.JTextField();
        jTextFieldXlsAddress = new javax.swing.JTextField();
        jButtonSaveXls = new javax.swing.JButton();
        jButtonSaveXlsx = new javax.swing.JButton();
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

        jButtonSaveDoc.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSaveDoc.setText("Экспортировать в .doc");
        jButtonSaveDoc.setToolTipText(null);
        jButtonSaveDoc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveDocActionPerformed(evt);
            }
        });
        jPanel1.add(jButtonSaveDoc);
        jButtonSaveDoc.setBounds(20, 230, 260, 80);

        jButtonSaveDocx.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSaveDocx.setText("Экспортировать в .docx");
        jButtonSaveDocx.setToolTipText(null);
        jButtonSaveDocx.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveDocxActionPerformed(evt);
            }
        });
        jPanel1.add(jButtonSaveDocx);
        jButtonSaveDocx.setBounds(20, 320, 260, 80);

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

        jButtonSaveXls.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSaveXls.setText("Экспортировать в .xls");
        jButtonSaveXls.setToolTipText(null);
        jButtonSaveXls.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveXlsActionPerformed(evt);
            }
        });
        jPanel2.add(jButtonSaveXls);
        jButtonSaveXls.setBounds(20, 230, 260, 80);

        jButtonSaveXlsx.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSaveXlsx.setText("Экспортировать в .xlsx");
        jButtonSaveXlsx.setToolTipText(null);
        jButtonSaveXlsx.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveXlsxActionPerformed(evt);
            }
        });
        jPanel2.add(jButtonSaveXlsx);
        jButtonSaveXlsx.setBounds(20, 320, 260, 80);

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

    private void jButtonSaveDocActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveDocActionPerformed
        logger.debug("Export to .doc button has been pressed");

        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        Export.toDoc(jTextFieldNumber.getText(), jTextFieldName.getText(), jTextFieldAddress.getText());

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveDocActionPerformed

    private void jTextFieldAddressActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextFieldAddressActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextFieldAddressActionPerformed

    private void jButtonSaveXlsActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveXlsActionPerformed
        logger.debug("Export to .xls button has been pressed");

        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        Export.toXls(
                jTextFieldXlsReceiverFIO.getText(), jTextFieldXlsName.getText(),
                jTextFieldXlsAddress.getText(), jTextFieldXlsSum.getText(),
                jTextFieldXlsSumUsl.getText());

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveXlsActionPerformed

    private void jButtonSaveDocxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveDocxActionPerformed
        logger.debug("Export to .docx button has been pressed");

        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        Export.toDocx(jTextFieldNumber.getText(), jTextFieldName.getText(), jTextFieldAddress.getText());

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveDocxActionPerformed

    private void jButtonSaveXlsxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveXlsxActionPerformed
        logger.debug("Export to .xlsx button has been pressed");

        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        Export.toXlsx(
                jTextFieldXlsReceiverFIO.getText(), jTextFieldXlsName.getText(),
                jTextFieldXlsAddress.getText(), jTextFieldXlsSum.getText(),
                jTextFieldXlsSumUsl.getText());

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveXlsxActionPerformed

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
    private javax.swing.JButton jButtonSaveDoc;
    private javax.swing.JButton jButtonSaveDocx;
    private javax.swing.JButton jButtonSaveXls;
    private javax.swing.JButton jButtonSaveXlsx;
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
