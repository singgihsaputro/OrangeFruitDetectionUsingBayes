/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.range;

import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel._Worksheet;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;

/**
 * This sample demonstrates how to protect rows and columns
 * from deleting but its content can be cleared.
 * Sample requires jexcel-full.jar in classpath
 *
 * @author Igor Novikov
 */
public class ColumnRowProtection
{
    public static void main(String[] args) throws Exception
    {
        final Application application = new Application();
        if (!application.isVisible())
        {
            application.setVisible(false);
        }

        final JWorkbook _workbook = new JWorkbook();
        try
        {
            _workbook.addWorksheet("probe");
        }
        catch (ExcelException e)
        {
            e.printStackTrace();
        }

        final Worksheet sheet = _workbook.getWorksheet("probe");

        final JFrame frame = new JFrame();
        frame.setContentPane(_workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 800);
        frame.setVisible(true);

        _workbook.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                _Worksheet _worksheet = sheet.getPeer();
                _worksheet.getCells().setLocked(new Variant(false));
                _worksheet.protect(
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        new Variant(true),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        new Variant(false),
                        new Variant(false),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter(),
                        Variant.createUnspecifiedParameter()
                );
            }
        });
    }
}