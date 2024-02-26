package org.shinear;
import javax.swing.*;
import java.awt.*;

public class EmailConverterApp {
    public static void main(String[] args) {
        JFrame window = new JFrame("Email Converter");
        window.setSize(600, 200);
        window.setLocation(100, 100);
        window.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        window.setLayout(new GridLayout(3, 1));

        // 创建一个面板对象，用于放置第一行的组件
        JPanel panel1 = new JPanel();
        // 设置面板的布局管理器
        panel1.setLayout(new FlowLayout(FlowLayout.LEFT));
        // 创建一个标签对象，用于显示“Open Email File:”
        JLabel label1 = new JLabel("Open Email File:");
        // 创建一个文本框对象，用于显示选择的原始电子邮件文件的路径
        JTextField textField1 = new JTextField(20);
        // 创建一个按钮对象，用于打开文件选择框
        JButton openButton = new JButton("Browse...");
        // 为按钮添加点击事件的监听器
        openButton.addActionListener(e -> {
            // 当按钮被点击时，打开文件选择框
            JFileChooser emailFileChooser = new JFileChooser();
            // 设置文件选择框的标题
            emailFileChooser.setDialogTitle("Select Email File");
            // 设置文件选择框的过滤器，只允许选择 .eml 格式的文件
            emailFileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
                @Override
                public boolean accept(java.io.File f) {
                    return f.isDirectory() || f.getName().endsWith(".eml");
                }

                @Override
                public String getDescription() {
                    return "Email Files (*.eml)";
                }
            });
            int result = emailFileChooser.showOpenDialog(window);
            // 如果用户选择了一个文件，将文件的路径显示在文本框中
            if (result == JFileChooser.APPROVE_OPTION) {
                java.io.File file = emailFileChooser.getSelectedFile();
                textField1.setText(file.getAbsolutePath());
            }
        });
        // 将标签、文本框和按钮添加到面板中
        panel1.add(label1);
        panel1.add(textField1);
        panel1.add(openButton);

        // 创建一个面板对象，用于放置第二行的组件
        JPanel panel2 = new JPanel();
        // 设置面板的布局管理器
        panel2.setLayout(new FlowLayout(FlowLayout.LEFT));
        // 创建一个标签对象，用于显示“Save Excel File:”
        JLabel label2 = new JLabel("Save Excel File:");
        // 创建一个文本框对象，用于显示选择的目标 Excel 文件的路径
        JTextField textField2 = new JTextField(20);
        // 创建一个按钮对象，用于打开文件选择框
        JButton targetButton = new JButton("Browse...");
        // 为按钮添加点击事件的监听器
        targetButton.addActionListener(e -> {
            // 当按钮被点击时，打开文件选择框
            JFileChooser directoryChooser = new JFileChooser();

            // 设置文件选择框的标题
            directoryChooser.setDialogTitle("Select Directory");
            // 设置文件选择框的模式，只允许选择目录
            directoryChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int result = directoryChooser.showSaveDialog(window);
            if (result == JFileChooser.APPROVE_OPTION) {
                java.io.File file = directoryChooser.getSelectedFile();
                // 检查目录是否存在
                if (file.exists()) {
                    textField2.setText(file.getAbsolutePath());
                } else {
                    JOptionPane.showMessageDialog(window, "Directory does not exist.");
                }
            }

        });
        // 将标签、文本框和按钮添加到面板中
        panel2.add(label2);
        panel2.add(textField2);
        panel2.add(targetButton);

        // 创建一个面板对象，用于放置第三行的组件
        JPanel panel3 = new JPanel();
        // 设置面板的布局管理器
        panel3.setLayout(new FlowLayout(FlowLayout.CENTER));
        // 创建一个按钮对象，用于转换邮件内容到 Excel 文档
        JButton convertButton  = new JButton("Convert to Excel");
        // 设置按钮的大小
        convertButton.setPreferredSize(new Dimension(200, 20));
        // 为按钮添加点击事件的监听器
        convertButton.addActionListener(e -> {
            // 当按钮被点击时，获取用户选择的原始电子邮件文件和目标 Excel 文件的路径
            String emailFilePath = textField1.getText();
            String excelFilePath = textField2.getText();
            // 如果用户没有选择文件，弹出提示框
            if (emailFilePath.isEmpty() || excelFilePath.isEmpty()) {
                JOptionPane.showMessageDialog(window, "Please select email file and excel file first.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
                // 执行转换操作
            Utils.convertEmailToExcel(emailFilePath, excelFilePath);
                            // 显示转换成功信息
            JOptionPane.showMessageDialog(window, "转换完成", "提示", JOptionPane.INFORMATION_MESSAGE);

        });

        // 将按钮添加到面板中
        panel3.add(convertButton);

        // 将三个面板添加到窗口中
        window.add(panel1);
        window.add(panel2);
        window.add(panel3);
        // 设置窗口可见
        window.setVisible(true);
    }


}