package msword;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;

public class ReceiptWord extends javax.swing.JFrame {

    Runnable saveTask;

    /**
     * Creates new form ReceiptWord
     */
    public ReceiptWord() {
        initComponents();

        this.saveTask = () -> {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");

            // Чтение из шаблона в переменную doc
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "receipt_template.doc")) {
                doc = new HWPFDocument(fis);
            } catch (Exception ex) {
                ex.printStackTrace();
                System.err.println("Error template!");
            }

            // Замена в переменной doc данных
            try {
                doc.getRange().replaceText("$ФИОплательщика", jTextFieldName.getText());
                doc.getRange().replaceText("$АДРЕСплательщика", jTextFieldAddress.getText());
            } catch (Exception ex) {
                System.err.println("Error replaceText!");
            }

            // Сохранение переменной doc в новый файл
            try (FileOutputStream fos = new FileOutputStream(dir + "receipt.doc")) {
                doc.write(fos);
                // Открытие файла внешней программой
                Desktop.getDesktop().open(new File(dir + "receipt.doc"));
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
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

        jTextFieldName = new javax.swing.JTextField();
        jTextFieldAddress = new javax.swing.JTextField();
        jButtonSave = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Квитанция MS Word");
        setResizable(false);
        getContentPane().setLayout(null);
        getContentPane().add(jTextFieldName);
        jTextFieldName.setBounds(500, 270, 530, 25);

        jTextFieldAddress.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextFieldAddressActionPerformed(evt);
            }
        });
        getContentPane().add(jTextFieldAddress);
        jTextFieldAddress.setBounds(500, 300, 530, 25);

        jButtonSave.setFont(new java.awt.Font("Tahoma", 1, 18)); // NOI18N
        jButtonSave.setText("Сохранить в Word");
        jButtonSave.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSaveActionPerformed(evt);
            }
        });
        getContentPane().add(jButtonSave);
        jButtonSave.setBounds(20, 260, 260, 80);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/msword/receipt.png"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(4, 4, 1040, 430);

        setBounds(0, 0, 1056, 478);
    }// </editor-fold>//GEN-END:initComponents

    private void jButtonSaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSaveActionPerformed
        // Устанавливаем иконку курсора в "загрузка"
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        // Запускаем сохранение в файл в отдельном потоке
        new Thread(saveTask).start();

        // Восстанавливаем стандартный курсор
        setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    }//GEN-LAST:event_jButtonSaveActionPerformed

    private void jTextFieldAddressActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextFieldAddressActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextFieldAddressActionPerformed

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
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextFieldAddress;
    private javax.swing.JTextField jTextFieldName;
    // End of variables declaration//GEN-END:variables
}
