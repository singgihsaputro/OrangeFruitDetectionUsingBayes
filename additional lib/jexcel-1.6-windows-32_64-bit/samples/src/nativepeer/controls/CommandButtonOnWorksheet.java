/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.controls;

import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.Automation;
import com.jniwrapper.win32.automation.AutomationException;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.ExcepInfo;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel.Shape;
import com.jniwrapper.win32.excel.Shapes;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;
import com.jniwrapper.win32.vbide._CodeModule;
import com.jniwrapper.win32.vbide._VBComponent;
import com.jniwrapper.win32.vbide._VBComponents;
import com.jniwrapper.win32.vbide._VBProject;

import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

/**
 * This sample demonstrates how to add command button to worksheet
 * and modify button properties using automation.
 * Sample requires jexcel-full.jar in classpath.
 */
public class CommandButtonOnWorksheet
{
    public static void main(String[] args) throws Exception
    {

        final JWorkbook workbook = new JWorkbook();
        try
        {
            workbook.addWorksheet("test");
        }
        catch (ExcelException e)
        {
            e.printStackTrace();
        }


        final Worksheet sheet = workbook.getWorksheet("test");

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


        sheet.getApplication().getOleMessageLoop().doInvokeLater(new Runnable()
        {
            public void run()
            {
                final int x = 140;
                final int y = 18;
                final int width = 100;
                final int height = 25;

                final Shapes shapes = sheet.getPeer().getShapes();
                final Variant unspecifiedParameter = Variant.createUnspecifiedParameter();
                final Shape testButton = shapes.addOLEObject(
                        new Variant("Forms.CommandButton.1"),
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        new Variant(x), new Variant(y),
                        new Variant(width), new Variant(height));

                shapes.setAutoDelete(false);
                shapes.release();

                System.out.println("Command button name before: " + testButton.getName());
                try
                {
                    testButton.setName(new BStr("Test"));
                    System.out.println("Command button name after: " + testButton.getName());
                    IDispatch oleObject = sheet.getPeer().OLEObjects(new Variant("Test"), new Int32(0));

                    Automation oleObjectAutomation = new Automation(oleObject, true);
                    IDispatch object = oleObjectAutomation.getProperty("Object").getPdispVal();
                    if (!object.isNull())
                    {
                        Automation intObjAutomation = new Automation(object, true);
                        intObjAutomation.setProperty("Caption", "Click Here");
                        intObjAutomation.setProperty("Enabled", "True");
                        intObjAutomation.release();

                        final String wbFunction = "Private Sub Test_Click()\n" +
                                "    Worksheets(1).Range(\"A1\").Value = 3.14159\n" +
                                "End Sub";

                        _VBProject vbProject = workbook.getWorkbook().getNativePeer().getVBProject();
                        _VBComponents components = vbProject.getVBComponents();
                        _VBComponent component = components.item(new Variant(1));
                        _CodeModule codeModule = component.getCodeModule();

                        codeModule.addFromString(new BStr(wbFunction));

                        codeModule.setAutoDelete(false);
                        codeModule.release();

                        component.setAutoDelete(false);
                        component.release();

                        components.setAutoDelete(false);
                        components.release();

                        vbProject.setAutoDelete(false);
                        vbProject.release();
                    }

                    object.release();
                    oleObjectAutomation.release();

                    oleObject.setAutoDelete(false);
                    oleObject.release();

                }
                catch (Exception ex)
                {
                    if (ex instanceof AutomationException)
                    {
                        AutomationException ex1 = (AutomationException) ex;
                        ExcepInfo exceptionInformation = ex1.getExceptionInformation();
                        String bstrDescription = exceptionInformation.getBstrDescription();
                        System.out.println("bstrDescription = " + bstrDescription);
                    }
                    ex.printStackTrace();
                }
            }
        });
    }
}