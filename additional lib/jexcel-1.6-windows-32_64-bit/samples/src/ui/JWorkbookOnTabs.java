/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package ui;

import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.ole.types.OleVerbs;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.util.TimerTask;

/**
 * This sample demonstrates embedded Excel workbooks on tabs.
 *
 * @author Igor Novikov
 */
public class JWorkbookOnTabs
{
    public static void main(String[] args) throws Exception
    {
        final JFrame frame = new JFrame("JWorkbook Tabs");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLocationRelativeTo(null);
        final Container cp = frame.getContentPane();
        cp.setLayout(new BorderLayout());

        final JTabbedPane tabbedPane = new JTabbedPane();
        cp.add(tabbedPane, BorderLayout.CENTER);

        final JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        cp.add(panel, BorderLayout.SOUTH);

        final JButton createButton = new JButton("Create JWorkbook tab");
        final JButton closeButton = new JButton("Close active tab");

        createButton.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent ae)
            {
                try
                {
                    System.out.println("OnClick");
                    final JWorkbook jworkbook = new JWorkbook();
                    tabbedPane.addTab("JWorkbook tab", jworkbook);
                    tabbedPane.setSelectedComponent(jworkbook);

                    closeButton.setEnabled(true);
                }
                catch (ExcelException e)
                {
                    e.printStackTrace();
                }
            }
        });

        closeButton.setEnabled(false);
        panel.add(createButton);
        panel.add(closeButton);

        closeButton.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent ae)
            {
                JWorkbook jworkbook = (JWorkbook) tabbedPane.getSelectedComponent();
                if (jworkbook != null)
                {
                    jworkbook.close();
                    tabbedPane.remove(jworkbook);
                }

                final int tabs = tabbedPane.getTabCount();
                if (tabs == 0) {
                    closeButton.setEnabled(false);
                }

            }
        });

        frame.setVisible(true);
    }
}