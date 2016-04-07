/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.*;

import java.util.List;

/**
 * This sample demonstrates how to listen to workbook and worksheet events.
 *
 * The sample works with MS Excel in non-embedded mode.
 *
 * @author Vladimir Kondrashchenko
 */
public class EventListenersSample
{
    public static void main(String[] args) throws ExcelException
    {
        //Start MS Excel application. Application starts invisible and without any workbooks
        Application application = new Application();
        Workbook workbook = application.createWorkbook(null);

        addListeners(workbook);

        runEventsGenerator(workbook);

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        workbook.close(false);
        application.close();
    }

    /**
     * Add event listeners to workbook and then iterate over its worksheets and add listeners to
     * each of them.
     *
     * @param workbook - workbook to add listeners to.
     */
    private static void addListeners(GenericWorkbook workbook)
    {
        WorkbookEventListenerImpl workbookEventListener = new WorkbookEventListenerImpl();
        WorksheetEventListenerImpl worksheetEventListener = new WorksheetEventListenerImpl();

        workbook.addWorkbookEventListener(workbookEventListener);

        List worksheets = workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            worksheet.addWorksheetEventListener(worksheetEventListener);
        }
    }

    /**
     * Perform actions on the workbook to demonstrate listeners' work
     *
     * @param workbook - workbook to perform actions
     * @throws ExcelException - thrown in case if addition of worksheet fails for some reasons
     */
    private static void runEventsGenerator(GenericWorkbook workbook) throws ExcelException
    {
        //Activate worksheets in turn
        List worksheets = workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            worksheet.activate();
        }

        //Create new worksheet
        Worksheet worksheet = workbook.addWorksheet("Custom sheet");
        worksheet.addWorksheetEventListener(new WorksheetEventListenerImpl());

        //Fill in cells
        for (int i = 1; i < 3; i++)
            for (int j = 1; j < 3; j++)
            {
                Cell cell = worksheet.getCell(i, j);
                cell.setValue(i * j);
            }

        //Clear filled cells
        worksheet.getRange("A1:D3").clear();
    }

    public static void log(String message)
    {
        System.out.println(message);
    }

    public static class WorkbookEventListenerImpl extends WorkbookEventAdapter
    {
        public void activate(WorkbookEventObject eventObject)
        {
            log("\"" + eventObject.getWorkbook().getWorkbookName() + "\" is activated.");
        }

        public void deactivate(WorkbookEventObject eventObject)
        {
            log("\"" + eventObject.getWorkbook().getWorkbookName() + "\" is deactivated.");
        }

        public void newSheet(WorkbookEventObject eventObject)
        {
            log("\"" + eventObject.getWorksheet().getName() + "\" is added to \"" +
                    eventObject.getWorkbook().getWorkbookName() + "\".");
        }

        public void beforeClose(WorkbookEventObject eventObject)
        {
            log("\"" + eventObject.getWorkbook().getWorkbookName() + "\" is closed.");
        }
    }

    public static class WorksheetEventListenerImpl extends WorksheetEventAdapter
    {
        public void changed(WorksheetEventObject eventObject)
        {
            log(eventObject.getRange().getAddress() + " is changed.");
        }

        public void activated(WorksheetEventObject eventObject)
        {
            log("\"" + eventObject.getWorksheet().getName() + "\" is activated.");
        }

        public void deactivated(WorksheetEventObject eventObject)
        {
            log("\"" + eventObject.getWorksheet().getName() + "\" is deactivated.");
        }
    }

}