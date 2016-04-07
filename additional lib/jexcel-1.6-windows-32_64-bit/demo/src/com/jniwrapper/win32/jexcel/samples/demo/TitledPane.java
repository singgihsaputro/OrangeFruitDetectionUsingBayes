/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package com.jniwrapper.win32.jexcel.samples.demo;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.AbstractBorder;
import javax.swing.event.PopupMenuEvent;
import javax.swing.event.PopupMenuListener;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

class TitledPane extends JPanel
{
    private long _hideTime = System.currentTimeMillis();
    private static final long HIDE_INTERVAL = 300;
    protected JLabel _title;
    protected JLabel _menuIcon;
    protected JLabel _closeIcon;

    TitledPane()
    {
        setLayout(new BorderLayout());

        final JPanel titlePanel = new JPanel();
        titlePanel.setLayout(new GridBagLayout());
        titlePanel.setBorder(null);

        _closeIcon = new JLabel();
        final Cursor handCursor = new Cursor(Cursor.HAND_CURSOR);
        _closeIcon.setCursor(handCursor);
        _closeIcon.setIcon(new ImageIcon(this.getClass().getResource("res/Close.gif")));
        _closeIcon.setToolTipText("Close");
        final Border emptyBorder = BorderFactory.createEmptyBorder(1, 2, 1, 2);
        final Border etchedBorder = BorderFactory.createEtchedBorder();
        _closeIcon.setBorder(emptyBorder);
        _closeIcon.addMouseListener(new MouseAdapter()
        {
            public void mouseEntered(MouseEvent e)
            {
                _closeIcon.setBorder(etchedBorder);
            }

            public void mouseExited(MouseEvent e)
            {
                _closeIcon.setBorder(emptyBorder);
            }
        });

        _menuIcon = new JLabel();
        _menuIcon.setCursor(handCursor);
        _menuIcon.setIcon(new ImageIcon(this.getClass().getResource("res/downArrow.gif")));
        _menuIcon.setToolTipText("Menu");
        _menuIcon.setBorder(emptyBorder);
        _menuIcon.addMouseListener(new MouseAdapter()
        {
            public void mouseEntered(MouseEvent e)
            {
                _menuIcon.setBorder(etchedBorder);
            }

            public void mouseExited(MouseEvent e)
            {
                _menuIcon.setBorder(emptyBorder);
            }
        });


        _title = new JLabel("");
        _title.setFocusable(false);
        _title.setBorder(null);

        final Font oldFont = _title.getFont();
        final Font newFont = new Font("Tahoma", Font.BOLD, oldFont.getSize());
        _title.setFont(newFont);

        titlePanel.add(_title, new GridBagConstraints(0, 0, 1, 1, 0.0, 0.0
                , GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(0, 5, 0, 0), 0, 0));

        titlePanel.add(_menuIcon, new GridBagConstraints(1, 0, 1, 1, 1.0, 0.0
                , GridBagConstraints.EAST, GridBagConstraints.NONE, new Insets(0, 0, 0, 0), 0, 0));

        titlePanel.add(_closeIcon, new GridBagConstraints(2, 0, 1, 1, 0.0, 0.0
                , GridBagConstraints.EAST, GridBagConstraints.NONE, new Insets(0, 5, 0, 5), 0, 0));

        titlePanel.add(new LineBevel(), new GridBagConstraints(0, 1, 3, 1, 1.0, 0.0
                , GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL, new Insets(0, 0, 0, 0), 0, 0));

        add(titlePanel, BorderLayout.NORTH);
    }


    TitledPane(String name,
               JComponent content,
               JPopupMenu actions,
               Action closeAction)
    {
        this();
        setTitle(name);
        setComponent(content);
        if (actions != null)
        {
            assignActions(actions);
        }
        if (closeAction != null)
        {
            assignCloseAction(closeAction);
        }
    }

    public void setTitle(String value)
    {
        _title.setText(value);
    }

    public void setComponent(JComponent component)
    {
        add(component, BorderLayout.CENTER);
    }

    public void assignActions(final JPopupMenu actions)
    {
        _menuIcon.addMouseListener(new MouseAdapter()
        {
            public void mouseClicked(MouseEvent e)
            {
                if ((System.currentTimeMillis() - _hideTime) < HIDE_INTERVAL)
                    return;

                actions.show(_menuIcon, 0, _menuIcon.getHeight());
            }
        });

        actions.addPopupMenuListener(new PopupMenuListener()
        {
            public void popupMenuCanceled(PopupMenuEvent e)
            {
            }

            public void popupMenuWillBecomeInvisible(PopupMenuEvent e)
            {
                _hideTime = System.currentTimeMillis();
            }

            public void popupMenuWillBecomeVisible(PopupMenuEvent e)
            {
            }
        });
    }

    public void assignCloseAction(final Action closeAction)
    {
        _closeIcon.addMouseListener(new MouseAdapter()
        {
            public void mouseClicked(MouseEvent e)
            {
                closeAction.actionPerformed(null);
            }
        });
    }

    private static class LineBorder extends AbstractBorder
    {
        public static final int HORISONTAL = 0;
        public static final int VERTICAL = 1;

        private int _type = HORISONTAL;

        LineBorder()
        {
        }

        LineBorder(int type)
        {
            _type = type;
        }

        public void paintBorder(Component c, Graphics g, int x, int y, int width, int height)
        {
            g.translate(x, y);

            if (_type == HORISONTAL)
            {
                int w = width;
                int h = height - 1;

                g.setColor(Color.lightGray);
                g.drawLine(0, h - 1, w - 1, h - 1);

                g.setColor(Color.white);
                g.drawLine(1, h, w, h);
            }
            else if (_type == VERTICAL)
            {
                int h = height;

                g.setColor(Color.lightGray);
                g.drawLine(0, 0, 0, h - 1);

                g.setColor(Color.white);
                g.drawLine(1, 1, 1, h);
            }

            g.translate(-x, -y);
        }

        public Insets getBorderInsets(Component c)
        {
            return new Insets(0, 0, 0, 0);
        }

        public boolean isBorderOpaque()
        {
            return true;
        }
    }

    private static class LineBevel extends JPanel
    {
        LineBevel(int type)
        {
            init(type);
        }

        private void init(int type)
        {
            setBorder(new LineBorder(type));

            if (type == LineBorder.HORISONTAL)
                setPreferredSize(new Dimension(1, 2));
            else
                setPreferredSize(new Dimension(2, 1));
        }

        LineBevel()
        {
            init(LineBorder.HORISONTAL);
        }
    }
}