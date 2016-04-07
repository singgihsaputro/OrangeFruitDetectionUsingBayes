/*
 * Copyright (c) 2000-2012 TeamDev Ltd. All rights reserved.
 * Use is subject to license terms.
 */

package basics;

import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.ExcelException;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.Window;

import java.io.IOException;
import java.util.List;

/**
 * This sample demonstrates how to obtain and modify workbook window settings. The sample will
 * work only with MS Application instance stated by itself and will not affect other MS Excel
 * applications running.
 *
 * The sample works with MS Excel in non-embedded mode.
 *
 * @author Vladimir Kondrashchenko
 */
public class WindowsSample
{
    public static void main(String[] args) throws ExcelException, IOException
    {
        //Start MS Excel application, crate workbook and make it visible.
        // Application starts invisible and without any workbooks
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);

        application.setVisible(true);

        printAllWindows(application);

        Window window = workbook.getWindow();

        System.out.println("Window properties before modifying:");
        printWindowProperties(window);

        waitUserReaction();

        modifyWindow(window);

        System.out.println("\nWindow properties after modifying:");
        printWindowProperties(window);

        waitUserReaction();

        //Perform cleanup after yourself and close the MS Excel application forcing it to quit
        application.close(true);
    }

    /**
     * Pause markers to allow user to see changes made to MS Excel windows
     */
    private static void waitUserReaction() throws IOException
    {
        System.out.println("\nHit enter to continue>");
        System.in.read();
    }

    /**
     * Print all workbook windows in scope of one particular
     * MS Excel application instance.
     * @param application - MS Excel application instance to list windows
     */
    public static void printAllWindows(Application application)
    {
        System.out.println("Windows list:");
        List windows = application.getWindows();
        for (int i = 0; i < windows.size(); i++)
        {
            Window window = (Window) windows.get(i);
            System.out.println('\t' + window.getCaption());
        }
        System.out.println();
    }

    /**
     * Print properties of the particular workbook window
     * @param window - window instance which properties we are going to print
     */
    public static void printWindowProperties(Window window)
    {
        System.out.println("Caption: " + window.getCaption());
        System.out.println("Width: " + window.getWidth());
        System.out.println("Height: " + window.getHeight());
        System.out.println("State: " + window.getState());
        System.out.println("Zoom: " + window.getZoom());
        System.out.println("Index: " + window.getIndex());
    }

    /**
     * Modify workbook window properties
     * @param window - workbook window which properties we are going to modify.
     */
    public static void modifyWindow(Window window)
    {
        window.setCaption("New window caption");
        window.setState(Window.State.NORMAL);
        window.setHeight(window.getHeight()/2);
        window.setWidth(window.getWidth()/2);
        window.setZoom(window.getZoom()*2);
    }
}