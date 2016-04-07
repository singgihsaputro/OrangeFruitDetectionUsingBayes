/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.controls;

import com.jniwrapper.win32.automation.Automation;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.automation.types.VariantBool;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.office.CommandBar;
import com.jniwrapper.win32.office.CommandBarControl;
import com.jniwrapper.win32.office.CommandBarControls;
import com.jniwrapper.win32.office.MsoControlType;
import com.jniwrapper.win32.office.impl.CommandBarControlsImpl;

import javax.swing.*;

/**
 * This sample creates custom menu.
 * Sample requires jexcel-full.jar in classpath.
 *
 * @author Igor Novikov
 */
public class ExcelCustomMenu
{
    public static void main(String[] args) throws Exception
    {

        final JWorkbook _workbook = new JWorkbook();
        try
        {
            _workbook.addWorksheet("test");
        }
        catch (ExcelException e)
        {
            e.printStackTrace();
        }


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
                Variant unspecifiedParameter = Variant.createUnspecifiedParameter();

                // Create a custom menu bar and set as the main application menu bar
                CommandBar menuBar = _workbook.getApplication().getPeer().getCommandBars().add(
                        unspecifiedParameter, unspecifiedParameter,
                        new Variant(true), new Variant(true));
                menuBar.setVisible(VariantBool.TRUE);

                // Add custom menu
                CommandBarControl testMenu = menuBar.getControls().add(
                        new Variant(MsoControlType.msoControlPopup),
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        new Variant(true));
                testMenu.setCaption(new BStr("&New Menu"));

                // Get internal Controls from popup menu using Automation
                Automation menuAutomation = new Automation(testMenu, true);
                IDispatch cbControls = menuAutomation.getProperty("Controls").getPdispVal();
                CommandBarControls commandBarControls = new CommandBarControlsImpl(cbControls);

                CommandBarControl menuItem;
                // Add five new menu items
                for (int i = 0; i < 5; i++)
                {
                    menuItem = commandBarControls.add(
                            new Variant(MsoControlType.msoControlButton),
                            unspecifiedParameter,
                            unspecifiedParameter,
                            unspecifiedParameter,
                            new Variant(true));
                    menuItem.setCaption(new BStr("MenuItem &" + (i + 1)));
                }

                menuAutomation.release();
            }
        });
    }
}