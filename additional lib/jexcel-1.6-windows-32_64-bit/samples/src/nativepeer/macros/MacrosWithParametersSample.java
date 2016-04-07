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
import com.jniwrapper.win32.vbide.*;

import java.io.File;

/**
 * This is modified MacroSample example to demonstrate way how
 * ot use VB Macros with parameters in JExcel
 * see also http://www.cpearson.com/excel/vbe.htm for some tips
 */
public class MacrosWithParametersSample
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
            workbook.getOleMessageLoop().doInvokeAndWait(new Runnable()
            {
                public void run()
                {
                    _VBProject vbProject = _workbook.getVBProject();

                    _VBComponents components = vbProject.getVBComponents();

                    vbProject.setAutoDelete(false);
                    vbProject.release();

                    _VBComponent component = components.item(new Variant(1));
                    String func = "Sub Sample_Func(cellName As Variant, cellValue As Variant)\n" +
                            "    Worksheets(1).Range(cellName).Value = cellValue\n" +
                            "End Sub";
                    _CodeModule codeModule = component.getCodeModule();

                    addProcedure(codeModule, func);

                    codeModule.setAutoDelete(false);
                    codeModule.release();

                    component.setAutoDelete(false);
                    component.release();

                    runProcedure(application,
                            workbook,
                            new Variant[]{
                                    new Variant(256.8567),
                                    new Variant("B4"),
                                new Variant("ThisWorkbook.Sample_Func")});
                    printProcedures(components);

                    components.setAutoDelete(false);
                    components.release();
                }
            });

            //move worksheet modified by macros to new workbook
            Worksheet srcWorksheet = workbook.getWorksheet(1);
            try {
                srcWorksheet.setName("Imported worksheet");
                workbook2.moveWorksheet(srcWorksheet, null);
                //save workbook
                workbook2.saveCopyAs(new File("D:\\test_workbook.xls"));
            } finally {
                srcWorksheet.release();
            }
        }
        finally
        {//following construction should be used to close excel even after exception was thrown
            workbook2.close(false);
            workbook2.release();
            workbook.close(false);
            workbook.release();

            application.close();
        }
    }

    /**
     * Invokes procedure
     *
     * @param app - application instance
     * @param workbook - workbook
     * @param procName - VBA procedure name
     */
    public static void runProcedure(Application app, Workbook workbook, String procName)
    {
        GenericWorkbook oldActiveWorkbook = app.getActiveWorkbook();
        if (!oldActiveWorkbook.equals(workbook))
        {
            workbook.activate();
        }
        Automation automation = new Automation(workbook.getApplication().getPeer(), true);
        try
        {
            automation.invoke("run", new Variant[]{new Variant(procName)});
            automation.release();
        }
        catch (AutomationException e)
        {
            String bstrDescription = e.getExceptionInformation().getBstrDescription();
            System.out.println("bstrDescription = " + bstrDescription);
            e.printStackTrace();
        }
        oldActiveWorkbook.activate();
        ((Workbook) oldActiveWorkbook).release();
    }

    /**
     * Invokes procedure
     *
     * @param app - application instance
     * @param workbook - workbook
     * @param params   - invocation params
     */
    public static void runProcedure(Application app, Workbook workbook, Variant[] params)
    {
        GenericWorkbook oldActiveWorkbook = app.getActiveWorkbook();
        if (!oldActiveWorkbook.equals(workbook))
        {
            workbook.activate();
        }
        Automation automation = new Automation(workbook.getApplication().getPeer(), true);
        try
        {
            automation.invoke("run", params);
            automation.release();
        }
        catch (AutomationException e)
        {
            e.printStackTrace();
        }
        oldActiveWorkbook.activate();
        ((Workbook) oldActiveWorkbook).release();
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

            component.setAutoDelete(false);
            component.release();

            System.out.println("module: " + module.getName());
            Int32 startLine = new Int32((int) module.getCountOfDeclarationLines().getValue() + 1);
            while (startLine.getValue() < module.getCountOfLines().getValue())
            {
                BStr procOfLine = module.getProcOfLine(startLine, new vbext_ProcKind(vbext_ProcKind.vbext_pk_Proc));
                System.out.println("\tprocedure: " + procOfLine);
                startLine = new Int32((int) (startLine.getValue() + module.getProcCountLines(procOfLine, new vbext_ProcKind(vbext_ProcKind.vbext_pk_Proc)).getValue()));
            }
            module.setAutoDelete(false);
            module.release();
        }
    }
}