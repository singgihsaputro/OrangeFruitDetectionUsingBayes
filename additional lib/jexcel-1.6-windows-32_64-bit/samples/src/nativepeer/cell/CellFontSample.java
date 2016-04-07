/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.cell;

import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel.Font;
import com.jniwrapper.win32.jexcel.Cell;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

/**
 * This sample demonstrates font decoration for cell content.
 * Sample requires jexcel-full.jar in classpath
 */
public class CellFontSample
{
    public static void main(String[] args) throws Exception {
        final JWorkbook _workbook = new JWorkbook();
        _workbook.addWorksheet("test");

        final Worksheet sheet = _workbook.getWorksheet("test");
        Cell cell = sheet.getCell("A1");
        cell.setValue("Something");

        final JFrame frame = new JFrame();
        frame.setContentPane(_workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 800);
        frame.setLocationRelativeTo(null);
        frame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                _workbook.close();
            }
        });
        frame.setVisible(true);

        sheet.getApplication().getOleMessageLoop().doInvokeAndWait(new Runnable() {
            public void run() {
                Cell cell = sheet.getCell("A1");
                // Set "Wrap text" option on
                cell.getPeer().setWrapText(new Variant(true));

                Font font = cell.getPeer().getCharacters(new Variant(1), new Variant(4)).getFont();
                font.setUnderline(new Variant(true));
                font.setBold(new Variant(true));
            }
        });
    }
}