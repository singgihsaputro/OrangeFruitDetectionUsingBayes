/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.macros;

import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.Automation;
import com.jniwrapper.win32.automation.AutomationException;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel._Workbook;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.Workbook;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.vbide._CodeModule;
import com.jniwrapper.win32.vbide._VBComponent;
import com.jniwrapper.win32.vbide._VBComponents;
import com.jniwrapper.win32.vbide.vbext_ProcKind;

import java.io.File;

/**
 * Date: May 3, 2006
 * Time: 3:21:22 PM
 * This sample demonstrates way to use VB Macros in JExcel
 * see also http://www.cpearson.com/excel/vbe.htm for some tips
 */
public class MacrosSample
{
    public static void main(String[] args) throws Exception
    {
        final Application application = new Application();
        final Workbook workbook = application.createWorkbook("Workbook1");
        final Workbook workbook2 = application.createWorkbook("Workbook2");
        try
        {
            //creating and executing macro in Workbook1
            final _Workbook _workbook = workbook.getNativePeer();
            workbook.getOleMessageLoop().doInvokeLater(new Runnable()
            {
                public void run()
                {
                    _VBComponents components = _workbook.getVBProject().getVBComponents();
                    components = (_VBComponents) workbook.getOleMessageLoop().bindObject(components);

                    _VBComponent component = components.item(new Variant(1));
                    String func = "Sub Sample_Func()\n" +
                            "    Worksheets(1).Range(\"A1\").Value = 3.14159\n" +
                            "End Sub";
                    addProcedure(component.getCodeModule(), func);
                    runProcedure(workbook, "ThisWorkbook.Sample_Func");
                    printProcedures(components);
                }
            });

            //move worksheet modified by macros to new workbook
            Worksheet srcWorksheet = workbook.getWorksheet(1);
            srcWorksheet.setName("Imported worksheet");
            workbook2.moveWorksheet(srcWorksheet, null);
            //save workbook
            workbook2.saveCopyAs(new File("d:\\test_workbook.xls"));

        }
        finally
        {//following construction should be used to close excel even after exception was thrown
            workbook2.close(false);
            workbook.close(false);

            application.close();
        }
    }

    /**
     * Invokes procedure
     *
     * @param workbook - workbook with VBA procedure
     * @param procName - VBA procedure name
     */
    public static void runProcedure(Workbook workbook, String procName)
    {
        Application app = workbook.getApplication();
        GenericWorkbook oldActiveWorkbook = app.getActiveWorkbook();
        if (!oldActiveWorkbook.equals(workbook))
        {
            workbook.activate();
        }
        Automation automation = new Automation(workbook.getApplication().getPeer(), true);
        try
        {
            automation.invoke("run", new Variant[]{new Variant(procName)});
        }
        catch (AutomationException e)
        {
            String bstrDescription = e.getExceptionInformation().getBstrDescription();
            System.out.println("bstrDescription = " + bstrDescription);
            e.printStackTrace();
        }
        oldActiveWorkbook.activate();
    }

    /**
     * Adds procedure to the given module
     *
     * @param module     - code module
     * @param procString - string representation of VBA procedure
     */
    public static void addProcedure(_CodeModule module, String procString)
    {
        Int32 count = new Int32((int) module.getCountOfLines().getValue() + 1);
        module.insertLines(count, new BStr(procString));
    }

    /**
     * Adds procedure that will be fired after given event occurs
     *
     * @param module     - code module
     * @param procString - string representation of VBA procedure
     * @param event      - string representation of VBA event
     */
    public static void addProcedure(_CodeModule module, String procString, String event)
    {
        Int32 line = module.createEventProc(new BStr(event), module.getName());
        module.insertLines(new Int32((int) line.getValue() + 1), new BStr(procString));
    }

    /**
     * Prints available procedures
     *
     * @param components - VBA components
     */
    public static void printProcedures(_VBComponents components)
    {
        for (long i = 1; i <= components.getCount().getValue(); i++)
        {
            _VBComponent component = components.item(new Variant(i));
            _CodeModule module = component.getCodeModule();
            System.out.println("module: " + module.getName());
            Int32 startLine = new Int32((int) module.getCountOfDeclarationLines().getValue() + 1);
            while (startLine.getValue() < module.getCountOfLines().getValue())
            {
                BStr procOfLine = module.getProcOfLine(startLine, new vbext_ProcKind(vbext_ProcKind.vbext_pk_Proc));
                System.out.println("\tprocedure: " + procOfLine);
                startLine = new Int32((int) (startLine.getValue() + module.getProcCountLines(procOfLine, new vbext_ProcKind(vbext_ProcKind.vbext_pk_Proc)).getValue()));
            }
        }
    }
}