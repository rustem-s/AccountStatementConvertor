package ru.mobilewin.accountstatementconvertor;

import com.intellij.uiDesigner.core.GridConstraints;
import com.intellij.uiDesigner.core.GridLayoutManager;
import com.intellij.uiDesigner.core.Spacer;
import org.apache.commons.io.FilenameUtils;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;

/**
 * Created by Rustem.Saidaliyev on 30.03.2016.
 */
public class Convertor extends JFrame {
    private JTextArea filesList;
    private JPanel panel1;
    private JButton chooseFiles;
    private JButton convert;
    private JButton clear;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(
                new Runnable() {
                    public void run() {
                        new Convertor().setupUI();
                    }
                }
        );

    }


    private void setupUI() {
        panel1 = new JPanel();
        panel1.setLayout(new GridLayoutManager(2, 5, new Insets(0, 0, 0, 0), -1, -1));
        filesList = new JTextArea();
        panel1.add(filesList, new GridConstraints(0, 0, 1, 5, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_WANT_GROW, null, new Dimension(150, 50), null, 0, false));
        final Spacer spacer1 = new Spacer();
        panel1.add(spacer1, new GridConstraints(1, 4, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_VERTICAL, 1, GridConstraints.SIZEPOLICY_WANT_GROW, null, null, null, 0, false));
        chooseFiles = new JButton();
        chooseFiles.setText("Выбрать файлы");
        panel1.add(chooseFiles, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        clear = new JButton();
        clear.setText("Очистить");
        panel1.add(clear, new GridConstraints(1, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        convert = new JButton();

        convert.setText("Конвертировать");
        panel1.add(convert, new GridConstraints(1, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));


        chooseFiles.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent e) {
                chooseFilesMouseClicked(e);
            }
        });


        clear.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent e) {
                clearMouseClicked(e);
            }
        });

        convert.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent e) {
                convertMouseClicked(e);
            }
        });

        setContentPane(panel1);
        setSize(500, 400);
        setLocationRelativeTo(null);
        setVisible(true);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        filesList.setText("");
        refreshButtons();
    }

    private void chooseFilesMouseClicked(ActionEvent e) {

//        String workingDir = System.getProperty("user.dir");
        String workingDir = "D:\\Win\\Convertor\\file\\in\\src";
        JFileChooser fileChooser = new JFileChooser(workingDir);
        FileFilter filter = new FileNameExtensionFilter("Excel workbook", "xlsx", "xlsb");
        fileChooser.addChoosableFileFilter(filter);
        fileChooser.setFileFilter(filter);

        fileChooser.setMultiSelectionEnabled(true);
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {

            File[] selectedFiles = fileChooser.getSelectedFiles();

            java.util.List<String> existingFiles = getFilesFromTextArea();

            for (File file : selectedFiles) {

                String fullName = file.getAbsolutePath();

                for (String existingFile : existingFiles) {
                    if (fullName.equalsIgnoreCase(existingFile)) {
                        JOptionPane.showMessageDialog(this, "Файл существует");
                        return;
                    }
                }

                filesList.append(String.format("%s\n", fullName));
            }


        }
        refreshButtons();
    }

    private void clearMouseClicked(ActionEvent e) {
        filesList.setText("");
        refreshButtons();
    }


    private void convertMouseClicked(ActionEvent e) {
        java.util.List<String> files = getFilesFromTextArea();

        java.util.List<String> convertedFiles = new ArrayList<String>();

        for (String file : files) {
            try {

                int convertFile = 0; // 0 - конвертировать, 2 - не конвертировать

                String fileWithoutExt = FilenameUtils.removeExtension(file);


                String newFile = fileWithoutExt + ".txt";

                File f = new File(newFile);

                if (f.exists() && !f.isDirectory()) {
                    convertFile = Utils.okcancel(String.format("Файл %s существует, заменить?", newFile));
                }
                if (convertFile == 2) {
                    continue;
                }

                String convertedContent = Utils.convertFile(file);

                PrintWriter writer = new PrintWriter(newFile, "UNICODE");
                writer.print(convertedContent);
                writer.close();

                convertedFiles.add(newFile);


            } catch (Exception ex) {
                StringWriter errors = new StringWriter();
                ex.printStackTrace(new PrintWriter(errors));
                String error = errors.toString();

                JOptionPane.showMessageDialog(this, error);
                System.out.println(error);
                return;
            }

        }

        if (convertedFiles.size() > 0) {
            String message = "Сконвертированные файлы:";
            for (String convertedFile : convertedFiles) {
                message = message + "\n" + convertedFile;
            }
            JOptionPane.showMessageDialog(this, message);
        }

    }

    private void refreshButtons() {
        java.util.List<String> files = getFilesFromTextArea();
        if (files.size() > 0) {
            convert.setEnabled(true);
            clear.setEnabled(true);
        } else {
            convert.setEnabled(false);
            clear.setEnabled(false);
        }
    }


    private java.util.List<String> getFilesFromTextArea() {
        java.util.List<String> result = new ArrayList<String>();

        String[] files = filesList.getText().split("\\n");
        for (String file : files) {
            if (file.isEmpty()) {
                continue;
            }
            result.add(file);
        }
        return result;
    }

    {
// GUI initializer generated by IntelliJ IDEA GUI Designer
// >>> IMPORTANT!! <<<
// DO NOT EDIT OR ADD ANY CODE HERE!
        $$$setupUI$$$();
    }

    /**
     * Method generated by IntelliJ IDEA GUI Designer
     * >>> IMPORTANT!! <<<
     * DO NOT edit this method OR call it in your code!
     *
     * @noinspection ALL
     */
    private void $$$setupUI$$$() {
        panel1 = new JPanel();
        panel1.setLayout(new GridLayoutManager(2, 5, new Insets(0, 0, 0, 0), -1, -1));
        filesList = new JTextArea();
        panel1.add(filesList, new GridConstraints(0, 0, 1, 5, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_WANT_GROW, null, new Dimension(150, 50), null, 0, false));
        final Spacer spacer1 = new Spacer();
        panel1.add(spacer1, new GridConstraints(1, 4, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_VERTICAL, 1, GridConstraints.SIZEPOLICY_WANT_GROW, null, null, null, 0, false));
        chooseFiles = new JButton();
        chooseFiles.setText("Выбрать файлы");
        panel1.add(chooseFiles, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        clear = new JButton();
        clear.setText("Очистить");
        panel1.add(clear, new GridConstraints(1, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        convert = new JButton();
        convert.setText("Конвертировать");
        panel1.add(convert, new GridConstraints(1, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
    }

    /**
     * @noinspection ALL
     */
    public JComponent $$$getRootComponent$$$() {
        return panel1;
    }
}
