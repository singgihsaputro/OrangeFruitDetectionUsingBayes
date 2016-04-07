/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.controls;

import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.automation.types.VariantBool;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.office.CommandBar;
import com.jniwrapper.win32.office.CommandBarControl;
import com.jniwrapper.win32.office.CommandBarControls;
import com.jniwrapper.win32.office._CommandBars;

import javax.swing.*;

/**
 * This sample creates copy of Excel menu bar for JWorkbook component.
 * Sample requires jexcel-full.jar in classpath.
 *
 * @author Igor Novikov
 */
public class ExcelMenuCopy
{
    public static void main(String[] args) throws Exception
    {
        final Application application = new Application();
        if (!application.isVisible())
        {
            application.setVisible(false);
        }

        final JWorkbook _workbook = new JWorkbook();

        final JFrame frame = new JFrame();
        frame.setContentPane(_workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 800);
        frame.setVisible(true);

        _workbook.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                final _CommandBars bars = _workbook.getApplication().getPeer().getCommandBars();
                CommandBarControls controls = null;
                for (int i = 1; i < bars.getCount().toLong().intValue(); i++)
                {
                    CommandBar bar = bars.getItem(new Variant(i));
                    if (bar.getName().toString().equals("Worksheet Menu Bar"))
                    {
                        controls = bar.getControls();
                        break;
                    }
                }
                bars.release();

                CommandBar menuBar = _workbook.getApplication().getPeer().getCommandBars().add(
                        new Variant("NewMenuBar"), Variant.createUnspecifiedParameter(),
                        new Variant(true), new Variant(true));

                assert controls != null;
                for (int j = 1; j <= controls.getCount().toLong().intValue(); j++)
                {
                    CommandBarControl control = controls.getItem(new Variant(j));
                    menuBar.getControls().add(Variant.createUnspecifiedParameter(),
                            new Variant(control.getId().getValue()), Variant.createUnspecifiedParameter(),
                            Variant.createUnspecifiedParameter(), Variant.createUnspecifiedParameter());
                }

                menuBar.setVisible(VariantBool.TRUE);
            }
        });
    }

}