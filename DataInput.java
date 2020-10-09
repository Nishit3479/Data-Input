import java.util.*;
import java.io.*;

import javax.swing.*;
import java.awt.event.*;

import org.apache.poi.ss.usermodel.*;

public class DataInput implements ActionListener
{
	static File test;
	static Workbook wb;
	static Sheet sh;
	static FileInputStream fis;
	static FileOutputStream fos;
	static Row row;
	static Cell cell;
	static JFrame frame;
	static JLabel lb1, lb2, lb3, lb4;
	static JButton bt1, bt2;
	static JTextField tf1, tf2, tf3;
	static int i = 1, j = 0, c = 0;
	public static void main(String [] args) throws Exception
	{
		new DataInput();
	}
	DataInput()
	{
		frame = new JFrame("Data Input");
		frame.setBounds(780,350,500,300);
		frame.setLayout(null);
		frame.setDefaultCloseOperation(frame.EXIT_ON_CLOSE);
		
		lb1 = new JLabel("Name : ");
		lb1.setBounds(55,51,100,25);
		frame.add(lb1);
		
		tf1 = new JTextField();
		tf1.setBounds(160,51,200,25);
		frame.add(tf1);
		
		lb2 = new JLabel("Reg No : ");
		lb2.setBounds(55,77,100,25);
		frame.add(lb2);
		
		tf2 = new JTextField();
		tf2.setBounds(160,77,200,25);
		frame.add(tf2);
		
		lb3 = new JLabel("email-id : ");
		lb3.setBounds(55,103,100,25);
		frame.add(lb3);
		
		tf3 = new JTextField();
		tf3.setBounds(160,103,200,25);
		frame.add(tf3);
		
		lb4 = new JLabel("©Java-Eclipse");
		lb4.setBounds(385,225,100,25);
		frame.add(lb4);
		
		bt1 = new JButton("Reset");
		bt1.setBounds(120,150,77,25);
		frame.add(bt1);
		bt1.addActionListener(this);
		
		bt2 = new JButton("Submit");
		bt2.setBounds(250,150,77,25);
		frame.add(bt2);
		bt2.addActionListener(this);
		
		frame.setVisible(true);
	}
	public void actionPerformed(ActionEvent e)
	{
		String S = e.getActionCommand();
        String[] options = {"Yes", "No"};
		if(S.equals("Reset"))
		{
			tf1.setText(null);
			tf2.setText(null);
			tf3.setText(null);
		}
		if(S.equals("Submit"))
		{
			c = JOptionPane.showOptionDialog
			        (
			        frame,
			        "Are You Sure? You want to Submit", 
			        "Warning",            
			        JOptionPane.YES_NO_OPTION,
			        JOptionPane.WARNING_MESSAGE,
			        null,     
			        options,  
			        options[0] 
			        );
			j = 1;
			if(c == JOptionPane.YES_OPTION)
			{
				excel();
				tf1.setText(null);
				tf2.setText(null);
				tf3.setText(null);
			}
		}
	}
	public static void excel()
	{
		File test = new File("C:\\Users\\NISHIT VERMA\\Documents\\test.xlsx");
		try {
			fis = new FileInputStream(test);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			wb = WorkbookFactory.create(fis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sh = wb.getSheet("Sheet1");
		row = sh.createRow(0);
		cell = row.createCell(0);
		cell.setCellValue(lb1.getText());
		cell = row.createCell(1);
		cell.setCellValue(lb2.getText());
		cell = row.createCell(2);
		cell.setCellValue(lb3.getText());
		row = sh.createRow(i);
		cell = row.createCell(0);
		cell.setCellValue(tf1.getText());
		cell = row.createCell(1);
		cell.setCellValue(tf2.getText());
		cell = row.createCell(2);
		cell.setCellValue(tf3.getText());
		try {
			fos = new FileOutputStream(test);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			wb.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			fos.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally
		{
			i++;
		}
	}
}
