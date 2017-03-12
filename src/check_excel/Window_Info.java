/*
 * Пакет, в котором находятся классы для проверки документа
 */
package check_excel;

import java.awt.Color;
import java.io.File;
import java.io.FileFilter;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

/**
 * Класс, в котором есть графическая оболочка, взаимодействует с классом checking.java. 
 * Производится указание пути к файлу .xls с последующей его обработкой в лругом классе
 * 
 * @author Alexandr
 */
public class Window_Info extends javax.swing.JFrame {
    
        checking obh = new checking();
    /**
     * Creates new form WindowForm
     */
    public Window_Info() {
        
        initComponents();
        
    }

                      
    private void initComponents() {

        information = new javax.swing.JLabel();
        choose_box = new javax.swing.JButton();
        check = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setBackground(new java.awt.Color(21, 156, 20));
        setForeground(new java.awt.Color(34, 174, 49));
        setResizable(false);

        information.setBackground(new java.awt.Color(0, 255, 255));
        information.setText("<html><p><font color=\"#FF4500\" "+ "size=\"4\" face=\"Verdana\">Информация</font> </p>"
            + "<font size=\"4\"><em>" + "<br>Данная программа проверяет документ, в котором<br>содержатся данные для создания аккаунтов<br></em></font></html>");

        choose_box.setBackground(new java.awt.Color(0, 204, 204));
        choose_box.setText("Выберите документ для проверки");
        choose_box.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                choose_boxActionPerformed(evt);
            }
        });

        check.setText("Проверить");
        check.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                checkActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(57, 57, 57)
                        .addComponent(information, javax.swing.GroupLayout.PREFERRED_SIZE, 283, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(95, 95, 95)
                        .addComponent(choose_box)))
                .addContainerGap(44, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(check)
                .addGap(148, 148, 148))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(information, javax.swing.GroupLayout.PREFERRED_SIZE, 149, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(36, 36, 36)
                .addComponent(choose_box)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(check)
                .addContainerGap(47, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>                        
    public class myFileFilter extends javax.swing.filechooser.FileFilter {
      String ext,description;

    public String getDescription() {
        return description;
    }

    myFileFilter(String ext, String description) {
          this.ext = ext;
      }
     //В этом методе может быть любая проверка файла
      public boolean accept(File f) {
          if(f != null) {
              if(f.isDirectory()) {
                  return true;
              }

              return f.toString().endsWith(ext);
          }
          return false;
      }
}
    /**
     * Создание окна выбора файла
     * 
     * @param evt 
     */
    private void choose_boxActionPerformed(java.awt.event.ActionEvent evt) {                                           
        JFileChooser dialog = new JFileChooser();
        dialog.setFileSelectionMode(JFileChooser.FILES_ONLY);
        dialog.setApproveButtonText("Выбрать");//выбрать название для кнопки согласия
        dialog.setDialogTitle("Выберите файл");// выбрать название
        dialog.setDialogType(JFileChooser.OPEN_DIALOG);// выбрать тип диалога Open или Save
        dialog.setMultiSelectionEnabled(false); // Разрегить выбор нескольки файлов
        dialog.showOpenDialog(this);
        File file = dialog.getSelectedFile();
        String s1 = file.getAbsolutePath();
        if(s1.indexOf(".xls")!=-1 && s1.indexOf(".xlsx")==-1){
            obh.wave_to_doc = file.getAbsolutePath();
        }
        else{
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Файл не является документом в формате .xls</i>");
        }
    }                                          

    private void checkActionPerformed(java.awt.event.ActionEvent evt) {                                      
        if(obh.wave_to_doc.equals(null)||obh.wave_to_doc.equals("")){
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Файл не должен быть пустым </i>");
        }
        else{
            obh.messages();
            System.exit(0);
        }
    }                                     

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Window_Info.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Window_Info.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Window_Info.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Window_Info.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
            
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Window_Info().setVisible(true);
                 
            }
        });
    }

    // Variables declaration - do not modify                     
    private javax.swing.JButton check;
    private javax.swing.JButton choose_box;
    private javax.swing.JLabel information;
    // End of variables declaration                   
}
