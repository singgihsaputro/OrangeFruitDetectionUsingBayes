/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.controls;

import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.automation.types.VariantBool;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.office.CommandBar;
import com.jniwrapper.win32.office._CommandBars;

import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

/**
 * Sample hides standard JWorkbook command bars.
 * Sample requires jexcel-full.jar in classpath.
 */
public class HideCommandBars
{
    public static void main(String[] args) throws Exception
    {
        final JWorkbook workbook = new JWorkbook();

        final JFrame frame = new JFrame();
        frame.setContentPane(workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 800);
        frame.addWindowListener(new WindowAdapter()
        {
            public void windowClosing(WindowEvent e)
            {
                workbook.close();
            }
        });
        frame.setVisible(true);

        workbook.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                try
                {
                    _CommandBars menuBars = workbook.getApplication().getPeer().getCommandBars();
                    int count = (int) menuBars.getCount().getValue();
                    for (int i = 1; i < count; i++)
                    {
                        CommandBar commandBar = menuBars.getItem(new Variant(i));
                        if (commandBar.getEnabled().equals(VariantBool.TRUE) && commandBar.getName().toString().indexOf("JCP-") == -1)
                        {
                            commandBar.setEnabled(VariantBool.FALSE);
                        }
                        commandBar.setAutoDelete(false);
                        commandBar.release();
                    }
                    menuBars.setAutoDelete(false);
                    menuBars.release();
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                }
            }
        });

    }
}