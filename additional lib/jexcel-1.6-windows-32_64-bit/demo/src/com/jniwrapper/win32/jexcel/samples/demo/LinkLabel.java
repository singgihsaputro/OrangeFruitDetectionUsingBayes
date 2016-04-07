/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo;

import javax.swing.Action;
import javax.swing.Icon;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Graphics;
import java.awt.Rectangle;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

/**
 * @author Vladimir Kondrashchenko
 */
class LinkLabel extends JLabel
{
    private static final Color INACTIVE = new Color(80, 105, 230);
    private static final Color ACTIVE = new Color(97, 170, 243);

    LinkLabel(String text, Icon icon, final Action action)
    {
        super(text, icon, SwingConstants.LEFT);
        setForeground(INACTIVE);

        addMouseListener(new MouseAdapter()
        {
            public void mouseEntered(MouseEvent e)
            {
                setForeground(ACTIVE);
                setCursor(new Cursor(Cursor.HAND_CURSOR));
            }

            public void mouseExited(MouseEvent e)
            {
                setForeground(INACTIVE);
                setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
            }

            public void mouseClicked(MouseEvent e)
            {
                action.actionPerformed(null);
            }
        });
    }

    public void paint(Graphics g)
    {
        super.paint(g);
        Rectangle bounds = getBounds();
        g.setColor(getForeground());
        g.drawLine(20, bounds.height - 2, bounds.width, bounds.height - 2);
    }
}