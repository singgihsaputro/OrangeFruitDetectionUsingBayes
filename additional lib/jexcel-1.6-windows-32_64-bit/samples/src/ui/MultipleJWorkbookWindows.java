/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package ui;

import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.LinkedList;
import java.util.List;

/**
 * This sample demonstrates multiple JWorkbook windows.
 *
 * @author Igor Novikov
 */
public class MultipleJWorkbookWindows
{
    private int counter = 1;
    private int frameCounter = 1;
    private final List jworkbookList = new LinkedList();

    public MultipleJWorkbookWindows()
    {
        JFrame frame = new JFrame("JWorkbook Demo");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        Container cp = frame.getContentPane();
        cp.setLayout(new BorderLayout());
        JPanel panel = new JPanel();
        cp.add(panel);
        panel.setLayout(new BorderLayout());
        JButton button = new JButton("Create JWorkbook");
        panel.add(button, BorderLayout.SOUTH);
        frame.setBounds(100, 100, 500, 70);
        frame.setLocation(10, 10);
        frame.setVisible(true);

        button.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent ae)
            {
                try
                {
                    JExcelWindow jframe = new JExcelWindow("JWorkbook window [" + Integer.toString(frameCounter) + "]");
                    jframe.setVisible(true);
                    jframe.setLocation(100 + counter * 20, 100 + counter * 20);
                    counter++;
                    frameCounter++;
                    counter = (counter > 10) ? 1 : counter;
                    jworkbookList.add(jframe.getJWorkbook());

                }
                catch (ExcelException e)
                {
                    e.printStackTrace();
                }
            }
        });

        frame.addWindowListener(new WindowAdapter()
        {
            public void windowClosing(WindowEvent e)
            {
                while (!jworkbookList.isEmpty())
                {
                    JWorkbook workbook = (JWorkbook) jworkbookList.remove(0);
                    if (workbook != null && !workbook.isClosed())
                    {
                        workbook.close();
                    }
                }
            }
        });
    }

    public static void main(String[] args)
    {
        new MultipleJWorkbookWindows();
    }

    private class JExcelWindow extends JFrame
    {
        private JWorkbook workbook;

        JExcelWindow(String title) throws ExcelException
        {
            setTitle(title);
            setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            setSize(700, 500);
            setLocation(100, 100);
            workbook = new JWorkbook();
            Container cp = this.getContentPane();
            cp.setLayout(new BorderLayout());
            cp.add(workbook, BorderLayout.CENTER);
            workbook.setVisible(true);
        }

        public JWorkbook getJWorkbook()
        {
            return workbook;
        }
    }
}