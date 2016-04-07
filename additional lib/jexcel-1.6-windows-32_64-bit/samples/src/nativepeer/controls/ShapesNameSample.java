/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.controls;

import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel.Shape;
import com.jniwrapper.win32.excel.Shapes;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.io.File;

/**
 * This sample loads workbook with comboboxes on worksheet
 * and parses their names.
 * Sample requires jexcel-full.jar in classpath.
 *
 * @author Igor Novikov
 */
public class ShapesNameSample
{
    public static void main(String[] args) throws Exception
    {
        //Please pay attention for path to combo.xls file
        final JWorkbook _workbook = new JWorkbook(new File("c://combo.xls"));
        final Worksheet sheet = _workbook.getWorksheet("test");

        final JFrame frame = new JFrame();
        frame.setContentPane(_workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 800);
        frame.setVisible(true);


        sheet.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                Shapes shapes = sheet.getPeer().getShapes();
                int num = (int) shapes.getCount().getValue();
                for (int i = 0; i < num; i++)
                {
                    Shape shape = shapes.item(new Variant(i + 1));
                    if (shape != null)
                        try
                        {
                            String name = shape.getName().toString();
                            System.out.println("CtXL.findListControl() shape" + i + ": " + name);

                        }
                        catch (Exception e)
                        {
                            System.out.println(e);
                        }
                }
            }
        });

    }
}