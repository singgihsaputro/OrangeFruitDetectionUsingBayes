/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.*;

import java.io.IOException;
import java.util.List;

/**
 * This sample demonstrates how to use workbook and worksheet event handlers. The sample shows how to
 * allow/disallow certain actions with help of the event handlers.
 *
 * The sample works with MS Excel in non-embedded mode.
 *
 * @author Vladimir Kondrashchenko
 */
public class EventHandlersSample
{
    public static void main(String[] args) throws ExcelException, IOException
    {
        //Start MS Excel application. Application starts invisible and without any workbooks
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);

        setupEventHandlers(workbook);

         //Open MS Excel application window
         application.setVisible(true);

        System.out.println("Press <Enter> to close application...");

        System.in.read();

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        application.close(true);
    }

    /**
     * Set event handlers to workbook and then iterate over its worksheets and set handlers to
     * each of them.
     * @param workbook - workbook to apply handlers to.
     */
    private static void setupEventHandlers(GenericWorkbook workbook)
    {
        WorkbookEventHandler workbookEventHandler = new WorkbookHandler();
        WorksheetEventHandler worksheetEventHandler = new WorksheetHandler();

        workbook.setEventHandler(workbookEventHandler);

        List worksheets = workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            worksheet.setEventHandler(worksheetEventHandler);
        }
    }

    public static class WorkbookHandler implements WorkbookEventHandler
    {
        public boolean beforeClose(WorkbookEventObject source)
        {
            //Allow closing any workbooks
            return true;
        }

        public boolean beforeSave(WorkbookEventObject source)
        {
            //Forbid saving any workbooks
            return false;
        }
    }

    public static class WorksheetHandler implements WorksheetEventHandler
    {
        public boolean beforeDoubleClick(WorksheetEventObject eventObject)
        {
            //Forbid double-clicking on "A1" cell
            return !eventObject.getCell().equals(eventObject.getWorksheet().getCell("A1"));
        }

        public boolean beforeRightClick(WorksheetEventObject eventObject)
        {
            //Allow right-clicking only on "A1" cell
            return eventObject.getRange().equals(eventObject.getWorksheet().getRange("A1"));
        }
    }
}