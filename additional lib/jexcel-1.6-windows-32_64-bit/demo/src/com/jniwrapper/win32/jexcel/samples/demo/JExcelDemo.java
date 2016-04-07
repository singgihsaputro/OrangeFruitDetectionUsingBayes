/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo;

import com.jniwrapper.util.Logger;
import com.jniwrapper.win32.gdi.Icon;
import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.jexcel.ui.*;
import com.jniwrapper.win32.ui.MessageBox;
import com.jniwrapper.win32.ui.Wnd;
import com.jniwrapper.win32.ui.dialogs.OpenSaveFileDialog;
import com.jniwrapper.PlatformContext;
import com.sun.java.swing.plaf.windows.WindowsMenuBarUI;

import javax.swing.*;
import java.awt.*;
import java.awt.Window;
import java.awt.event.*;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

/**
 * @author Vladimir Kondrashchenko
 * @author Dmitriy Milashenko
 */

public class JExcelDemo extends JFrame
{
    private static final Logger LOG = Logger.getInstance(JExcelDemo.class);

    private static int framesCount = 0;

    private final int excelVersion;

    static final int VERSION_NUMBER_2007 = 12;
    private static final int DIVIDER = 2;
    private static final int AUTOMATION_BAR_WIDTH = 244;
    private static final String FILTER = "Excel Workbooks (*.xls) | *.xls";
    private static final String FILTER_12 = "Excel12 Workbooks (*.xlsx) | *.xlsx";
    public static final String DEMO_TITLE = "JExcel Demo";

    private final JWorkbook _workbook;

    private final AutomationBarAction _automationBarAction = new AutomationBarAction();
    private final EventLogAction _eventLogAction = new EventLogAction();
    private final StaticViewAction _staticViewAction = new StaticViewAction();

    private final WorksheetEventLogger _worksheetEventLogger = new WorksheetEventLogger();
    private final WorkbookEventLogger _workbookEventLogger = new WorkbookEventLogger();

    private final JWorkbookListener jworkbookEventListener = new JWorkbookListener();
    private final DemoPrintPreviewListener printPreviewListener = new DemoPrintPreviewListener();

    private JSplitPane _splitPane;

    private TitledPane _eventLogPane;
    private JTextArea _eventLog;

    private JSplitPane _subSplitPane;
    private TitledPane _automationBarPane;
    private AutomationPanel _automationPanel;

    private JMenuBar _menuBar;
    private JMenuItem _miSave;
    private JMenuItem _miSaveAs;
    private JMenuItem _miPrintPreview;
    private JMenuItem _miPrint;
    private JMenuItem _miNew;
    private JMenuItem _miOpen;

    private JCheckBoxMenuItem _miStaticView;

    public JExcelDemo() throws ExcelException
    {
        super(DEMO_TITLE + " - [New Workbook]");
        framesCount++;

        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        getContentPane().setLayout(new BorderLayout());
        //create jWorkbook
        _workbook = new JWorkbook();
        if (!_workbook.isWorkbookCreated())
            throw new IllegalStateException("Failed to obtain a workbook");

        Application app = _workbook.getApplication();
        excelVersion = (int) Float.parseFloat(app.getVersion(Application.LANGUAGE_NEUTRAL_LCID));

        _workbook.addJWorkbookEventListener(jworkbookEventListener);
        _workbook.addPrintPreviewListener(printPreviewListener);

        _splitPane = createSplitPane();
        getContentPane().add(_splitPane, BorderLayout.CENTER);

        _menuBar = createMenuTree();
        setJMenuBar(_menuBar);

        initUi();
        updateRecentFiles();
        attachEventLoggers();
        attachListeners();
    }

    private void initUi()
    {
        setSize(new Dimension(800, 600));
        setLocationRelativeTo(null);
    }

    private void attachListeners()
    {
        addWindowListener(new WindowAdapter()
        {
            public void windowOpened(WindowEvent e)
            {
                setupFrameIcon(JExcelDemo.this);
            }

            public void windowClosing(WindowEvent e)
            {
                closeWorkbook();
            }
        });
    }

    private void closeWorkbook()
    {
        _workbook.close();
        _workbook.removeJWorkbookEventListener(jworkbookEventListener);
        _workbook.removePrintPreviewListener(printPreviewListener);
    }

    private JSplitPane createSplitPane()
    {
        JSplitPane splitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, true);
        splitPane.setBorder(null);
        splitPane.setDividerSize(DIVIDER);
        splitPane.setResizeWeight(0.8);

        _subSplitPane = createSubSplitPane();
        splitPane.setTopComponent(_subSplitPane);

        _eventLogPane = createEventLogPane();
        splitPane.setBottomComponent(_eventLogPane);

        return splitPane;
    }

    private JSplitPane createSubSplitPane()
    {
        _automationPanel = new AutomationPanel(_workbook, this);
        _automationPanel.setMinimumSize(new Dimension(AUTOMATION_BAR_WIDTH, 473));
        _automationPanel.setMaximumSize(new Dimension(AUTOMATION_BAR_WIDTH, 0));
        _automationBarPane = new TitledPane("Automation Bar", _automationPanel, null, _automationBarAction);

        final JSplitPane subSplitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, true);
        subSplitPane.setLeftComponent(_workbook);
        subSplitPane.setRightComponent(_automationBarPane);
        subSplitPane.setBorder(null);
        subSplitPane.setDividerSize(DIVIDER);
        subSplitPane.addComponentListener(new ComponentAdapter()
        {
            public void componentResized(ComponentEvent e)
            {
                subSplitPane.setDividerLocation(subSplitPane.getWidth() - AUTOMATION_BAR_WIDTH - DIVIDER);
            }
        });
        subSplitPane.addPropertyChangeListener(JSplitPane.DIVIDER_LOCATION_PROPERTY, new PropertyChangeListener()
        {
            public void propertyChange(PropertyChangeEvent evt)
            {
                subSplitPane.setDividerLocation(subSplitPane.getWidth() - AUTOMATION_BAR_WIDTH - DIVIDER);
            }
        });
        return subSplitPane;
    }

    private TitledPane createEventLogPane()
    {
        JPanel logPanel = createLogPanel();
        JPopupMenu popupMenu = createPopupMenu();
        return new TitledPane("Event Log", logPanel, popupMenu, _eventLogAction);
    }

    private JPanel createLogPanel()
    {
        _eventLog = new JTextArea();
        _eventLog.setFont(new Font("Courier New", Font.PLAIN, 11));
        _eventLog.setEditable(false);
        _eventLog.setBorder(null);

        JScrollPane eventLogScrollPane = new JScrollPane(_eventLog);
        eventLogScrollPane.setBorder(null);

        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.add(eventLogScrollPane, BorderLayout.CENTER);
        logPanel.setMinimumSize(new Dimension(200, 80));

        return logPanel;
    }

    private JPopupMenu createPopupMenu()
    {
        JPopupMenu popupMenu = new JPopupMenu();
        popupMenu.add(new JMenuItem(new AbstractAction("Clear")
        {
            public void actionPerformed(ActionEvent e)
            {
                _eventLog.setText("");
            }
        }));
        return popupMenu;
    }

    private String getFileFilter()
    {
        if (getExcelVersion() < VERSION_NUMBER_2007)
            return FILTER;
        else
            return FILTER_12;
    }

    int getExcelVersion()
    {
        return excelVersion;
    }

    private void attachEventLoggers()
    {
        List worksheets = _workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            worksheet.addWorksheetEventListener(_worksheetEventLogger);
        }
        _workbook.addWorkbookEventListener(_workbookEventLogger);
    }

    private void removeEventLoggers()
    {
        List worksheets = _workbook.getWorksheets();
        for (int i = 0; i < worksheets.size(); i++)
        {
            Worksheet worksheet = (Worksheet) worksheets.get(i);
            worksheet.removeWorksheetEventListener(_worksheetEventLogger);
        }
        _workbook.removeWorkbookEventListener(_workbookEventLogger);
    }

    private void enableMenuItems()
    {
        _miPrint.setEnabled(true);
        _miPrintPreview.setEnabled(true);
        _miSave.setEnabled(false);
        _miSaveAs.setEnabled(true);
    }

    private void printLog(String s)
    {
        _eventLog.append(s + '\n');
        SwingUtilities.invokeLater(new Runnable()
        {
            public void run()
            {
                _eventLog.setCaretPosition(_eventLog.getText().length());
            }
        });
    }

    private JMenuBar createMenuTree()
    {
        JMenuBar menuBar = new JMenuBar();
        menuBar.setUI(new WindowsMenuBarUI());

        JMenu fileMenu = new JMenu("File");
        fileMenu.setMnemonic(KeyEvent.VK_F);

        _miNew = new JMenuItem(new AbstractAction("New...")
        {
            public void actionPerformed(ActionEvent e)
            {
                try
                {
                    final JExcelDemo excelDemo = new JExcelDemo();
                    excelDemo.addWindowListener(new WindowAdapter()
                    {
                        public void windowOpened(WindowEvent e)
                        {
                            setupFrameIcon(excelDemo);
                        }
                    });
                    excelDemo.setVisible(true);
                }
                catch (ExcelException e1)
                {
                    JOptionPane.showMessageDialog(null, "Cannot start MS Excel application!\n" +
                            "JExcel Demo requires MS Excel to be installed.",
                            "JExcel Demo", JOptionPane.WARNING_MESSAGE);
                }
//                _workbook.newWorkbook();
//                _miSave.setEnabled(false);
            }
        });
        _miNew.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_N, ActionEvent.CTRL_MASK));

        _miOpen = new JMenuItem(new AbstractAction("Open...")
        {
            public void actionPerformed(ActionEvent e)
            {
                OpenSaveFileDialog dialog = new OpenSaveFileDialog(JExcelDemo.this);
                dialog.setFilter(getFileFilter());

                if (dialog.getOpenFileName())
                {
                    try
                    {
                        _workbook.openWorkbook(new File(dialog.getFileName()));
                        _miSave.setEnabled(false);
                        if (_workbook.getFile() != null)
                        {
                            JExcelDemo.this.setTitle(DEMO_TITLE + " - [" + _workbook.getFile().getAbsolutePath() + "]");
                        }
                    }
                    catch (FileNotFoundException e1)
                    {
                        JOptionPane.showMessageDialog(SwingUtilities.windowForComponent(JExcelDemo.this),
                                "Cannot find file " + dialog.getFileName() + ".", "File not found",
                                JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });
        _miOpen.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, ActionEvent.CTRL_MASK));

        JMenuItem miClose = new JMenuItem(new AbstractAction("Close")
        {
            public void actionPerformed(ActionEvent e)
            {
                closeWorkbook();
                JExcelDemo.this.setVisible(false);
                JExcelDemo.this.dispose();
            }
        });

        _miSave = new JMenuItem(new AbstractAction("Save")
        {
            public void actionPerformed(ActionEvent e)
            {

                if (_workbook.getFile() != null)
                {
                    try
                    {
                        _workbook.save();
                    }
                    catch (IOException el)
                    {
                        JOptionPane.showMessageDialog(JExcelDemo.this,
                                "Unable to save the workbook to the specified file.",
                                "JExcel Demo",
                                JOptionPane.ERROR_MESSAGE);
                    }
                }
                else
                {
                    OpenSaveFileDialog dialog = new OpenSaveFileDialog(JExcelDemo.this);
                    dialog.setFilter(getFileFilter());

                    if (dialog.getSaveFileName())
                    {
                        try
                        {
                            _workbook.saveAs(new File(dialog.getFileName()), _workbook.getFileFormat(), false);
                            if (_workbook.getFile() != null)
                            {
                                JExcelDemo.this.setTitle(DEMO_TITLE + " - [" + _workbook.getFile().getAbsolutePath() + "]");
                            }
                        }
                        catch (IOException el)
                        {
                            JOptionPane.showMessageDialog(JExcelDemo.this,
                                    "Unable to save the workbook to the specified file.",
                                    "JExcel Demo",
                                    JOptionPane.ERROR_MESSAGE);
                        }
                    }
                }
            }
        });
        _miSave.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, ActionEvent.CTRL_MASK));

        _miSaveAs = new JMenuItem(new AbstractAction("Save As...")
        {
            public void actionPerformed(ActionEvent e)
            {
                OpenSaveFileDialog dialog = new OpenSaveFileDialog(JExcelDemo.this);
                dialog.setFilter(getFileFilter());

                if (dialog.getSaveFileName())
                {
                    try
                    {
                        _workbook.saveAs(new File(dialog.getFileName()), _workbook.getFileFormat(), false);
                        _miSave.setEnabled(true);
                        if (_workbook.getFile() != null)
                        {
                            JExcelDemo.this.setTitle(DEMO_TITLE + " - [" + _workbook.getFile().getAbsolutePath() + "]");
                        }
                    }
                    catch (IOException el)
                    {
                        JOptionPane.showMessageDialog(JExcelDemo.this,
                                "Unable to save the workbook to the specified file.",
                                "JExcel Demo",
                                JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });

        _miPrintPreview = new JMenuItem(new AbstractAction("Print Preview")
        {
            public void actionPerformed(ActionEvent e)
            {
                _workbook.setPrintPreview(true);
            }
        });

        _miPrint = new JMenuItem(new AbstractAction("Print...")
        {
            public void actionPerformed(ActionEvent e)
            {
                _workbook.showPrintDialog();
            }
        });
        _miPrint.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_P, ActionEvent.CTRL_MASK));

        JMenuItem miExit = new JMenuItem(new AbstractAction("Exit")
        {
            public void actionPerformed(ActionEvent e)
            {
                final boolean printPreview = _workbook.isPrintPreview();
                if (printPreview)
                {
                    _workbook.setPrintPreview(false);
                }
                else
                {
                    processWindowEvent(new WindowEvent(JExcelDemo.this, WindowEvent.WINDOW_CLOSING));
                }
            }
        });

        fileMenu.add(_miNew);
        fileMenu.add(_miOpen);
        fileMenu.add(miClose);
        fileMenu.addSeparator();
        fileMenu.add(_miSave);
        fileMenu.add(_miSaveAs);
        fileMenu.addSeparator();
        fileMenu.add(_miPrintPreview);
        fileMenu.add(_miPrint);
        fileMenu.addSeparator();
        fileMenu.addSeparator();
        fileMenu.add(miExit);
        fileMenu.addItemListener(new ItemListener()
        {
            public void itemStateChanged(ItemEvent e)
            {
                if (e.getStateChange() == ItemEvent.SELECTED)
                {
                    updateRecentFiles();
                }
            }
        });

        JMenu viewMenu = new JMenu("View");
        viewMenu.setMnemonic(KeyEvent.VK_V);

        JCheckBoxMenuItem miShowAutomationBar = new JCheckBoxMenuItem(_automationBarAction);
        _automationBarAction.setMenuItem(miShowAutomationBar);
        miShowAutomationBar.setSelected(true);
        JCheckBoxMenuItem miShowEventLog = new JCheckBoxMenuItem(_eventLogAction);
        _eventLogAction.setMenuItem(miShowEventLog);
        miShowEventLog.setSelected(true);

        _miStaticView = new JCheckBoxMenuItem(_staticViewAction);
        _staticViewAction.setMenuItem(_miStaticView);

        viewMenu.add(miShowAutomationBar);
        viewMenu.add(miShowEventLog);
        /*viewMenu.addSeparator();
        viewMenu.add(_miStaticView);*/

        JMenu helpMenu = new JMenu("Help");
        JMenuItem about = new JMenuItem(new ShowAboutAction(this));
        helpMenu.add(about);

        menuBar.add(fileMenu);
        menuBar.add(viewMenu);
        menuBar.add(helpMenu);

        return menuBar;
    }

    public void dispose()
    {
        super.dispose();

        framesCount--;
        if (framesCount == 0)
        {
            System.exit(0);
        }
    }

    private void updateRecentFiles()
    {
        if (!_workbook.isWorkbookCreated())
        {
            return;
        }

        JMenu fileMenu = _menuBar.getMenu(0);
        while (fileMenu.getItemCount() > 11)
        {
            fileMenu.remove(9);
        }

        List files;
        try
        {
            files = _workbook.getApplication().getRecentFiles();
        }
        catch (Exception e)
        {
            return;
        }

        for (int i = 0; i < files.size(); i++)
        {
            final File file = (File) files.get(i);
            final int maxLength = 40;
            String path = pathText(file, maxLength);
            RecentFilesAction action = new RecentFilesAction(path,
                    JExcelDemo.this,
                    _workbook,
                    file);
            JMenuItem recentFile = new JMenuItem(action);
            fileMenu.insert(recentFile, 9 + i);
        }

        _automationPanel.updateRecentFiles();
    }

    static String pathText(File file, int maxLength)
    {
        String test = file.getAbsolutePath();
        if (test.length() > maxLength)
        {
            int pos = test.lastIndexOf('\\');
            if (pos != -1 && (maxLength - test.length() + pos) > 0)
            {
                String subPath = test.substring(0, maxLength - test.length() + pos);
                test = test.substring(0, subPath.lastIndexOf('\\') + 1) + "..." + test.substring(pos, test.length());
            }
        }
        return test;
    }

    private void setMenuEnabled(boolean enabled)
    {
        JMenu fileMenu = _menuBar.getMenu(0);
        for (int i = 0; i < fileMenu.getItemCount() - 1; i++)
        {
            if (fileMenu.getItem(i) != null)
            {
                fileMenu.getItem(i).setEnabled(enabled);
            }
        }
    }


    private class JWorkbookListener extends JWorkbookEventAdapter
    {
        public void newWorkbook(JWorkbookEventObject eventObject)
        {
            updateUi();
            JExcelDemo.this.setTitle(DEMO_TITLE + " - [New Workbook]");
            printLog("New workbook: " + eventObject.getWorkbook().getWorkbookName() + ".");
        }

        public void workbookOpened(JWorkbookEventObject eventObject)
        {
            updateUi();
            if (_workbook.getFile() != null)
            {
                JExcelDemo.this.setTitle(DEMO_TITLE + " - [" + _workbook.getFile().getAbsolutePath() + "]");
            }
            else
            {
                JExcelDemo.this.setTitle(DEMO_TITLE);
            }
            printLog("Opened workbook: " + eventObject.getWorkbook().getWorkbookName() + ".");
        }

        private void updateUi()
        {
            _staticViewAction.setSelected(false);
            attachEventLoggers();
            updateRecentFiles();
            _automationPanel.updateRecentFiles();
            enableMenuItems();
        }

        public void beforeWorkbookClose(JWorkbookEventObject eventObject) throws JWorkbookInterruptException
        {
            if (!_workbook.isSaved())
            {
                int msgBoxresult = MessageBox.show(new Wnd(_workbook),
                        "Excel confirmation",
                        "Do you want to save changes you made to '" + _workbook.getWorkbookName() + "'",
                        MessageBox.YESNO);
                if (msgBoxresult == MessageBox.IDYES)
                {
                    openSaveDialog();
                }
            }

            removeEventLoggers();
            printLog("Workbook \"" + eventObject.getWorkbook().getWorkbookName() + "\" closed.");

            _staticViewAction.setSelected(false);
            _miPrint.setEnabled(false);
            _miPrintPreview.setEnabled(false);
            _miSave.setEnabled(false);
            _miSaveAs.setEnabled(false);
            _miStaticView.setEnabled(false);

            _automationPanel.setSelection(null);
        }

        private void openSaveDialog()
                throws JWorkbookInterruptException
        {
            Window parent = SwingUtilities.getWindowAncestor(_workbook);
            OpenSaveFileDialog saveDialog = new OpenSaveFileDialog(parent);
            saveDialog.setFilter(getFileFilter());
            boolean result = saveDialog.getSaveFileName();
            if (result)
            {
                String newFileName = saveDialog.getFileName();
                File newFile = new File(newFileName);
                try
                {
                    _workbook.saveCopyAs(newFile);
                }
                catch (IOException e)
                {
                    throw new JWorkbookInterruptException("Input/Output error", e);
                }
            }
        }
    }

    class EventLogAction extends AbstractAction
    {
        private boolean _selected;
        private JCheckBoxMenuItem _menuItem;

        EventLogAction()
        {
            super("Show Event Log");
            _selected = true;
        }

        public void setMenuItem(JCheckBoxMenuItem menuItem)
        {
            _menuItem = menuItem;
        }

        //show/hide evet log panel
        public void actionPerformed(ActionEvent e)
        {
            _selected = !_selected;
            _eventLogPane.setVisible(_selected);
            if (_selected)
            {
                _splitPane.setDividerLocation(0.8);
            }
            _menuItem.setSelected(_selected);
        }
    }

    class AutomationBarAction extends AbstractAction
    {
        private boolean _selected;
        private JCheckBoxMenuItem _menuItem;

        AutomationBarAction()
        {
            super("Show Automation Bar");
            _selected = true;
        }

        public void setMenuItem(JCheckBoxMenuItem menuItem)
        {
            _menuItem = menuItem;
        }

        //show/hide automation sidebar
        public void actionPerformed(ActionEvent e)
        {
            _selected = !_selected;
            _automationBarPane.setVisible(_selected);
            if (_selected)
            {
                _subSplitPane.setDividerLocation(_subSplitPane.getWidth() - AUTOMATION_BAR_WIDTH - DIVIDER);
            }
            _menuItem.setSelected(_selected);
        }
    }

    class StaticViewAction extends AbstractAction
    {
        private boolean _selected;
        private JCheckBoxMenuItem _menuItem;

        StaticViewAction()
        {
            super("Static view");
            _selected = false;
        }

        public void setMenuItem(JCheckBoxMenuItem menuItem)
        {
            _menuItem = menuItem;
        }

        //switch views between static and active
        public void actionPerformed(ActionEvent e)
        {
            _selected = !_selected;
            _menuItem.setSelected(_selected);
            _workbook.setStaticMode(_selected);
            enableMenuItems();
        }

        public void setSelected(boolean value)
        {
            _selected = value;
            _menuItem.setSelected(value);
            enableMenuItems();
        }

        private void enableMenuItems()
        {
            _miPrint.setEnabled(!_selected);
            _miPrintPreview.setEnabled(!_selected);
            _miSave.setEnabled(false);
            _miSaveAs.setEnabled(!_selected);
            _miNew.setEnabled(!_selected);
            _miOpen.setEnabled(!_selected);
        }
    }

    class ShowAboutAction extends AbstractAction
    {
        Window _owner;

        ShowAboutAction(Window owner)
        {
            super("About JExcel");
            _owner = owner;
        }

        public void actionPerformed(ActionEvent e)
        {
            AboutDialog aboutDialog = new AboutDialog((Frame) _owner);
            aboutDialog.setVisible(true);
        }
    }


    class WorksheetEventLogger extends WorksheetEventAdapter
    {
        public void activated(WorksheetEventObject e)
        {
            printLog(e.getWorksheet().getName() + " activated.");
        }

        public void beforeDoubleClick(WorksheetEventObject e)
        {
            printLog("Double click on " + e.getCell().getAddress() + ".");
        }

        public void beforeRightClick(WorksheetEventObject e)
        {
            printLog("Right click on " + e.getRange().getAddress() + ".");
        }

        public void changed(WorksheetEventObject e)
        {
            printLog(e.getRange().getAddress() + " is changed.");
        }

        public void deactivated(WorksheetEventObject e)
        {
            printLog(e.getWorksheet().getName() + " deactivated.");
        }

        public void sheetCalculated(WorksheetEventObject e)
        {
            printLog(e.getWorksheet().getName() + " calculated.");
        }

        public void selectionChanged(WorksheetEventObject eventObject)
        {
            Range range = eventObject.getRange();
            _automationPanel.setSelection(range);
            printLog(range.getAddress() + " selected.");
            range.setAutoDelete(false); // notify JExcel not to release this object, since it's used in _automationPanel
        }
    }

    private class DemoPrintPreviewListener extends PrintPreviewAdapter
    {
        public void onPrintPreview(final JWorkbookEventObject eventObject)
        {
            printLog("Print preview mode.");
            setMenuEnabled(false);
        }

        public void onPrintPreviewExit(final JWorkbookEventObject eventObject)
        {
            printLog("Print preview mode is closed.");
            setMenuEnabled(true);
        }
    }

    class WorkbookEventLogger extends WorkbookEventAdapter
    {
        public void newSheet(WorkbookEventObject eventObject)
        {
            Worksheet worksheet = eventObject.getWorksheet();
            printLog("\"" + worksheet.getName() + "\" sheet is added to \"" +
                    eventObject.getWorkbook().getWorkbookName() + "\" workbook.");
            worksheet.addWorksheetEventListener(_worksheetEventLogger);
            worksheet.setAutoDelete(false);
        }
    }

    private static void setupFrameIcon(java.awt.Window owner)
    {
        try
        {
            final Wnd winWnd = new Wnd(owner);
            final Icon bigIcon = new Icon(JExcelDemo.class.getResourceAsStream("res/jexcel.ico"), new Dimension(32, 32));
            final Icon smallIcon = new Icon(JExcelDemo.class.getResourceAsStream("res/jexcel.ico"), new Dimension(16, 16));
            winWnd.setWindowIcon(smallIcon, Icon.IconType.SMALL);
            winWnd.setWindowIcon(bigIcon, Icon.IconType.BIG);
        }
        catch (IOException e)
        {
            LOG.error("", e);
        }
    }


    private static void setupLookAndFeel()
    {
        String className = UIManager.getSystemLookAndFeelClassName();
        try
        {
            UIManager.setLookAndFeel(className);

            UIDefaults defaults = UIManager.getDefaults();
            Font tahoma = new Font("Tahoma", 0, 11);
            defaults.put("MenuItem.font", tahoma);
            defaults.put("TabbedPane.font", tahoma);
            defaults.put("TextField.font", tahoma);
            defaults.put("Label.font", tahoma);
            defaults.put("Button.font", tahoma);
            defaults.put("ComboBox.font", tahoma);
            defaults.put("Panel.font", tahoma);
        }
        catch (ClassNotFoundException e)
        {
            LOG.error("LookAndFeel class not found: " + className, e);
        }
        catch (IllegalAccessException e)
        {
            LOG.error("LookAndFeel class access failure: " + className, e);
        }
        catch (InstantiationException e)
        {
            LOG.error("Can not create instance of class: " + className, e);
        }
        catch (UnsupportedLookAndFeelException e)
        {
            LOG.error("Unsupported LookAndFeel class: " + className, e);
        }
    }

    static
    {
        setupLookAndFeel();
        JPopupMenu.setDefaultLightWeightPopupEnabled(false);
        JFrame.setDefaultLookAndFeelDecorated(true);
    }

    public static void main(String[] args)
    {
        if (!PlatformContext.isWindows())
        {
            JOptionPane.showMessageDialog(null, "JExcel Demo can be launch in MS Windows enviroment only!",
                    "JExcel Demo", JOptionPane.WARNING_MESSAGE);
        }

        try
        {
            final JExcelDemo excelDemo = new JExcelDemo();
            excelDemo.setVisible(true);
        }
        catch (ExcelException e)
        {
            JOptionPane.showMessageDialog(null, "Cannot start MS Excel application!\n" +
                    "JExcel Demo requires MS Excel to be installed.",
                    "JExcel Demo", JOptionPane.WARNING_MESSAGE);
        }
    }
}