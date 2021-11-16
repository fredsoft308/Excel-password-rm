package officepasswordremover;

import java.io.File;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

public class OfficePasswordRemove extends javax.swing.JFrame {

    private File selectedFile = null;

    public OfficePasswordRemove() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        fileChooserBtn = new javax.swing.JButton();
        fileTextField = new javax.swing.JTextField();
        convertBtn = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        fileChooserBtn.setText("Datei wählen");
        fileChooserBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileChooserBtnActionPerformed(evt);
            }
        });

        convertBtn.setText("Passwort entfernen");
        convertBtn.setEnabled(false);
        convertBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                convertBtnActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addComponent(fileTextField)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(fileChooserBtn))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(convertBtn)
                        .addGap(0, 232, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(fileChooserBtn)
                    .addComponent(fileTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(convertBtn)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void fileChooserBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_fileChooserBtnActionPerformed

        FileNameExtensionFilter filterWord = new FileNameExtensionFilter("Word-Dokumente", "docx");
        FileNameExtensionFilter filterWordTemplate = new FileNameExtensionFilter("Word-Vorlagen", "dotx");
        FileNameExtensionFilter filterExcel = new FileNameExtensionFilter("Excel-Dokumente", "xlsx");
        JFileChooser fileChooser = new JFileChooser();
        //fileChooser.setFileFilter(filter);
        fileChooser.addChoosableFileFilter(filterWord);
        fileChooser.addChoosableFileFilter(filterWordTemplate);
        fileChooser.addChoosableFileFilter(filterExcel);
        fileChooser.setAcceptAllFileFilterUsed(false);

        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        
        File select = this.selectedFile;
        if(select == null){
            select = new File(this.fileTextField.getText());
        }
        
        fileChooser.setCurrentDirectory(select);

        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {

            this.selectedFile = fileChooser.getSelectedFile();
            this.fileTextField.setText(this.selectedFile.getAbsolutePath());
            this.convertBtn.setEnabled(true);
        }
    }//GEN-LAST:event_fileChooserBtnActionPerformed

    private void convertBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_convertBtnActionPerformed
        this.convertBtn.setEnabled(false);

        if (!this.selectedFile.isDirectory()) {
            new OfficePasswordRemover(this.selectedFile);
        } else {
            File[] listOfFiles = this.selectedFile.listFiles();
            for (File file : listOfFiles) {
                if (file.isFile()) {
                    new OfficePasswordRemover(file);
                }
            }
        }
        this.convertBtn.setEnabled(true);
        JOptionPane pane = new JOptionPane("Passwort entfernt");
        JDialog dialog = pane.createDialog("OPR");
        dialog.setAlwaysOnTop(true);
        dialog.setVisible(true);

    }//GEN-LAST:event_convertBtnActionPerformed

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
            java.util.logging.Logger.getLogger(OfficePasswordRemove.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(OfficePasswordRemove.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(OfficePasswordRemove.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(OfficePasswordRemove.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new OfficePasswordRemove().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton convertBtn;
    private javax.swing.JButton fileChooserBtn;
    private javax.swing.JTextField fileTextField;
    // End of variables declaration//GEN-END:variables
}
