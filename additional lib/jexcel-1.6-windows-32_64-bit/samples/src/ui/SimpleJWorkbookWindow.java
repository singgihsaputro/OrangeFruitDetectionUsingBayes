/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package ui;

import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.jexcel.ui.JWorkbookEventAdapter;
import com.jniwrapper.win32.jexcel.ui.JWorkbookEventObject;
import com.jniwrapper.win32.jexcel.ui.JWorkbookInterruptException;
import com.jniwrapper.win32.ui.MessageBox;
import com.jniwrapper.win32.ui.Wnd;
import com.jniwrapper.win32.ui.dialogs.OpenSaveFileDialog;

import javax.swing.*;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;

/**
 * This sample demonstrates how to create simple window with integrated Excel workbook.
 */
public class SimpleJWorkbookWindow extends JFrame
{
    private JWorkbook uiWorkbook;

    public SimpleJWorkbookWindow()
    {
        super("Simple JWorkbook Window");

        //Create JWorkbook
        uiWorkbook = null;
        try
        {
            uiWorkbook = new JWorkbook();
        }
        catch (ExcelException e)
        {
            throw new RuntimeException("Unable to create JWorkbook", e);
        }

        uiWorkbook.addJWorkbookEventListener(new JWorkbookEventAdapter()
        {
            public void beforeWorkbookClose(JWorkbookEventObject eventObject) throws JWorkbookInterruptException
            {
                GenericWorkbook workbook = eventObject.getWorkbook();
                if (!workbook.isSaved())
                {
                    int msgBoxresult = MessageBox.show(new Wnd(SimpleJWorkbookWindow.this),
                            "Excel confirmation",
                            "Do you want to save changes you made to '" + workbook.getWorkbookName() + "'",
                            MessageBox.YESNO);
                    if (msgBoxresult == MessageBox.IDYES)
                    {
                        Window parent = SwingUtilities.getWindowAncestor(SimpleJWorkbookWindow.this);
                        OpenSaveFileDialog saveDialog = new OpenSaveFileDialog(parent);
                        String FILTER = "Excel Workbooks (*.xls) | *.xls";
                        saveDialog.setFilter(FILTER);
                        boolean result = saveDialog.getSaveFileName();
                        if (result)
                        {
                            String newFileName = saveDialog.getFileName();
                            File newFile = new File(newFileName);
                            try
                            {
                                workbook.saveCopyAs(newFile);
                            }
                            catch (IOException e)
                            {
                                throw new JWorkbookInterruptException("Input/Output error", e);
                            }
                        }
                    }
                }
            }
        });

        //Insert the JWorkbook into JFrame
        Container contentPane = getContentPane();
        contentPane.setLayout(new BorderLayout());
        contentPane.add(uiWorkbook, BorderLayout.CENTER);

        //Add window state listener to clean up after yourself
        addWindowListener(new WindowAdapter()
        {
            public void windowClosing(WindowEvent e)
            {
                uiWorkbook.close();
            }
        });
    }

    public static void main(String[] args)
    {
        SimpleJWorkbookWindow sampleWindow = new SimpleJWorkbookWindow();
        sampleWindow.setSize(800, 600);
        sampleWindow.setLocationRelativeTo(null);
        sampleWindow.setDefaultCloseOperation(EXIT_ON_CLOSE);
        sampleWindow.setVisible(true);
    }
}