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
import com.jniwrapper.win32.office.CommandBarControl;
import com.jniwrapper.win32.office.CommandBarControls;
import com.jniwrapper.win32.office._CommandBars;

import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

/**
 * This sample demonstrates how to disable a particular control from a menu bar using JExcel.
 */
public class DisableMenuControl
{
    public static void main(String[] args) throws Exception
    {
        final JWorkbook _workbook = new JWorkbook();

        final JFrame frame = new JFrame();
        frame.setContentPane(_workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLocationRelativeTo(null);

        frame.addWindowListener(new WindowAdapter()
        {
            public void windowClosing(WindowEvent e)
            {
                _workbook.close();
            }
        });

        _workbook.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                try
                {
                    _CommandBars menuBars = _workbook.getApplication().getPeer().getCommandBars();
                    CommandBar standard = menuBars.getItem(new Variant("Standard"));
                    CommandBarControls commandBarControls = standard.getControls();
                    CommandBarControl control = commandBarControls.getItem(new Variant("&Save"));
                    control.setEnabled(new VariantBool(false));
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                }
            }
        });

        frame.setVisible(true);
    }
}