/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ApplicationEventListener;
import com.jniwrapper.win32.jexcel.ApplicationEventObject;

/**
 * This sample demonstrates how to listen to MS Excel application events
 * such as opening and creating workbook.
 *
 * The sample works with MS Excel in non-embedded mode.
 *
 * WARNING: do not close MS Excel window, hit enter in the command line, otherwise sample will fail on exit
 *
 * @author Vladimir Kondrashchenko
 */
public class ApplicationEventsSample
{
    public static void main(String[] args) throws Exception
    {
        Application application = new Application();
        application.addApplicationEventListener(new ApplicationEventListener()
        {
            public void newWorkbook(ApplicationEventObject eventObject)
            {
                System.out.println(eventObject.getWorkbook().getWorkbookName() + " workbook is created.");
            }

            public void openWorkbook(ApplicationEventObject eventObject)
            {
                System.out.println(eventObject.getWorkbook().getWorkbookName() + " workbook is opened.");
            }
        });

        //Open MS Excel application window
        application.setVisible(true);

        System.out.println("Press <Enter> to close the application...");

        System.in.read();

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        application.close(true);
    }
}