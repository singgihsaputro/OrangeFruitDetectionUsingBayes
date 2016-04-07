/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package tkbayes;

/**
 *
 * @author sion
 */
import java.awt.image.BufferedImage;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;

public class rgb extends JFrame {

    int jmlred, jmlgreen, jmlblue;
    double ratared, ratagreen, ratablue;

    public BufferedImage imageGray(BufferedImage src) {
        BufferedImage dest = new BufferedImage(src.getWidth(), src.getHeight(),
                BufferedImage.TYPE_INT_RGB);

        for (int y = 0; y < src.getHeight(); y++) {
            for (int x = 0; x < src.getWidth(); x++) {
                int rgb = src.getRGB(x, y);

                int alpha = (rgb << 24) & 0xFF;
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = (rgb) & 0xFF;
                int avg = (red + green + blue) / 3;
                int gray = alpha | avg << 16 | avg << 8 | avg;

                dest.setRGB(x, y, gray);
                jmlred += red;
                jmlgreen += green;
                jmlblue += blue;

            }
        }
        ratared = (double) jmlred / (double) (src.getHeight() * src.getWidth());
        ratagreen = (double) jmlgreen / (double) (src.getHeight() * src.getWidth());
        ratablue = (double) jmlblue / (double) (src.getHeight() * src.getWidth());
        return dest;
    }

    public BufferedImage imageTracehold(BufferedImage src) {
        BufferedImage dest = new BufferedImage(src.getWidth(), src.getHeight(),
                BufferedImage.TYPE_INT_RGB);
        int gray = 0;

        for (int y = 0; y < src.getHeight(); y++) {
            for (int x = 0; x < src.getWidth(); x++) {
                int rgb = src.getRGB(x, y);

                // Merubah Warna Ke RGB
                int red = rgb & 0x000000FF;
                int green = (rgb & 0x0000FF00) >> 8;
                int blue = (rgb & 0x00FF0000) >> 16;
                // END OF Merubah Warna Ke RGB

                int avg = (red + green + blue) / 3;
                // Nilai 128 di bawah ini dapat Anda ubah sesuai dengan
                // kebutuhan
                if (avg < 230) {
                    gray = 0;
                } 
                else {
                    gray = 255;
                }

                // Merubah RGB Ke Warna
                int biner = gray + (gray << 8) + (gray << 16);
                // END OF Merubah RGB Ke Warna
                dest.setRGB(x, y, biner);
            }
        }
        return dest;
    }

    public BufferedImage imageMaxFilter(BufferedImage src) {
        BufferedImage dest = new BufferedImage(src.getWidth(), src.getHeight(),
                BufferedImage.TYPE_INT_RGB);
        int big;
        int data[][] = new int[src.getHeight()][src.getWidth()];
        for (int y = 0; y < src.getHeight(); y++) {
            for (int x = 0; x < src.getWidth(); x++) {
                data[y][x] = src.getRGB(x, y);
            }
        }
        for (int y = 0; y < src.getHeight() - 2; y++) {
            for (int x = 0; x < src.getWidth() - 2; x++) {
                big = 99999999;
                for (int q = y; q < y + 3; q++) {
                    for (int w = x; w < x + 3; w++) {
                        if (big > data[q][w]) {
                            big = data[q][w];
                        }
                    }
                }
                dest.setRGB(x + 1, y + 1, big);
            }
        }
        return dest;
    }

    public void tampilImage(BufferedImage src) {
        setSize(400, 150);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        try {
            ImageIcon ico = new ImageIcon(src);
            JLabel label = new JLabel(ico);
            label.setBounds(10, 10, src.getWidth(), src.getHeight());
            add(label);
        } catch (Exception e) {
        }
        setSize(500, 500);
        setLocationRelativeTo(null);
        setLayout(null);
        setVisible(true);
    }
    public int batasx, batasy, batasa, batasb, diameter;

    public void getDiameter(BufferedImage src) {
        int rgb = src.getRGB(100, 100);
        int temp1 = 0, temp2 = 0, y = 1, x = 1, a = src.getWidth() - 2, b = src.getHeight() - 2;
        while (temp1 != rgb && x < 250) {
            y = 1;
            while (temp1 != rgb && y < 195) {
                temp1 = src.getRGB(x, y);

                y++;
            }
            x++;
        }

        while (temp2 != rgb && a > 0) {
            b = src.getHeight() - 2;
            while (temp2 != rgb && b > 0) {
                temp2 = src.getRGB(a, b);
                b--;
            }
            a--;
        }
        batasx = x;
        batasy = y;
        batasa = a;
        batasb = b;
        diameter = (a - x);

    }
}
