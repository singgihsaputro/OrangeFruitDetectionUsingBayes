/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.util.Calendar;

/**
 * @author Vladimir Ikryanov
 */
public class AboutDialog extends JDialog {
    public AboutDialog(Frame owner) {
        super(owner, "About JExcel Demo", true);
        initContent();
        initKeyStroke();
        setResizable(false);
        pack();
        setSize(300, getHeight());
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
    }

    private void initKeyStroke() {
        JRootPane rootPane = getRootPane();
        KeyStroke keyStroke = KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0, false);
        rootPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(keyStroke, "ESCAPE");
        rootPane.getActionMap().put("ESCAPE", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                dispose();
            }
        });
    }

    private void initContent() {
        JLabel icon = new JLabel(new ImageIcon(getClass().getResource("res/logo.jpg")));
        JLabel appName = new JLabel("JExcel Demo");
        JLabel version = new JLabel("Version 1.6");
        JLabel company = new JLabel("\u00A9 " + Calendar.getInstance().get(Calendar.YEAR) + " TeamDev Ltd.");
        JLabel rights = new JLabel("All rights reserved.");

        icon.setAlignmentX(Component.CENTER_ALIGNMENT);
        appName.setAlignmentX(Component.CENTER_ALIGNMENT);
        appName.setFont(appName.getFont().deriveFont(Font.BOLD, 12.0f));
        version.setAlignmentX(Component.CENTER_ALIGNMENT);
        company.setAlignmentX(Component.CENTER_ALIGNMENT);
        rights.setAlignmentX(Component.CENTER_ALIGNMENT);

        JPanel contentPane = new JPanel();
        contentPane.setBackground(Color.WHITE);
        contentPane.setBorder(new EmptyBorder(10, 10, 10, 10));
        contentPane.setLayout(new BoxLayout(contentPane, BoxLayout.Y_AXIS));
        contentPane.add(icon);
        contentPane.add(Box.createVerticalStrut(16));
        contentPane.add(appName);
        contentPane.add(Box.createVerticalStrut(8));
        contentPane.add(version);
        contentPane.add(Box.createVerticalStrut(8));
        contentPane.add(company);
        contentPane.add(rights);
        setContentPane(contentPane);
    }
}