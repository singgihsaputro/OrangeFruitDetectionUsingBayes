package tkbayes;

/**
 *
 * @author sion
 */

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.util.ArrayList;

public class ImageHelper {

    private ArrayList koordObjek;
    private BufferedImage realImage;
    private BufferedImage grayImage;
    private BufferedImage binaryImage;
    private int tinggiCitra;
    private int lebarCitra;
    private double diameterObjek;
    private double meanR;
    private double meanG;
    private double meanB;

    public ImageHelper() {
        this.realImage = null;
        this.grayImage = null;
        this.binaryImage = null;
        this.meanR = 0;
        this.meanG = 0;
        this.meanB = 0;
        this.koordObjek = new ArrayList();
        this.diameterObjek = 0;
        this.lebarCitra = 0;
        this.tinggiCitra = 0;
    }

    public static BufferedImage resize(BufferedImage img, int newW, int newH) {
        int w = img.getWidth();
        int h = img.getHeight();

        BufferedImage dimg = dimg = new BufferedImage(newW, newH, img.getType());
        Graphics2D g = dimg.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g.drawImage(img, 0, 0, newW, newH, 0, 0, w, h, null);
        g.dispose();

        return dimg;
    }

    public void setImage(BufferedImage input) {
        this.realImage = resize(input, 400, 300);
        this.setSize();
        this.imageToGray();
        this.imageToBinary();
        this.hitungDiameter();
        this.hitungMeanRGB();
    }

    public BufferedImage getImage() {
        return this.realImage;
    }

    public BufferedImage getGrayImage() {
        return this.grayImage;
    }

    public BufferedImage getBinaryImage() {
        return this.binaryImage;
    }

    public double getDiameter() {
        return this.diameterObjek;
    }

    public double getR() {
        return this.meanR;
    }

    public double getG() {
        return this.meanG;
    }

    public double getB() {
        return this.meanB;
    }

    private void setSize() {
        this.tinggiCitra = realImage.getHeight();
        this.lebarCitra = realImage.getWidth();
    }

    private void imageToGray() {
        // Init variable
        double red, green, blue;
        int gray;
        Color before, after;

        BufferedImage output = new BufferedImage(lebarCitra, tinggiCitra,
                BufferedImage.TYPE_BYTE_GRAY);

        for (int y = 0; y < tinggiCitra; y++) {
            for (int x = 0; x < lebarCitra; x++) {

                before = new Color(realImage.getRGB(x, y) & 0x00ffffff);

                // Calculate RGB to gray
                // with lumonisity algorithm
                red = (double) (before.getRed() * 0.21);
                green = (double) (before.getGreen() * 0.71);
                blue = (double) (before.getBlue() * 0.07);
                gray = (int) (red + green + blue);

                after = new Color(gray, gray, gray);
                output.setRGB(x, y, after.getRGB());
            }
        }
        this.grayImage = output;
    }

    private void imageToBinary() {
        Color before, after;
        koordObjek.clear();
        BufferedImage output = new BufferedImage(lebarCitra, tinggiCitra,
                BufferedImage.TYPE_BYTE_GRAY);

        for (int y = 0; y < tinggiCitra; y++) {
            for (int x = 0; x < lebarCitra; x++) {

                before = new Color(this.grayImage.getRGB(x, y) & 0x00ffffff);
                if (before.getBlue() < 251) {
                    after = new Color(255, 255, 255);
                    koordObjek.add(String.valueOf(x) + "," + String.valueOf(y));
                } else {
                    after = new Color(0, 0, 0);
                }
                output.setRGB(x, y, after.getRGB());
            }
        }
        this.binaryImage = output;
    }

     private void hitungMeanRGB() {
        int length = koordObjek.size();
        double red = 0;
        double green = 0;
        double blue = 0;
        int x, y;
        Color before;
        String xy[];
        for (int i = 0; i < length; i++) {
            xy = koordObjek.get(i).toString().split(",");
            x = Integer.parseInt(xy[0]);
            y = Integer.parseInt(xy[1]);
            before = new Color(this.realImage.getRGB(x, y) & 0x00ffffff);
            red += before.getRed();
            blue += before.getBlue();
            green += before.getGreen();
        }
        red /= length;
        blue /= length;
        green /= length;

        this.meanR = red;
        this.meanG = green;
        this.meanB = blue;
    }

    private void hitungDiameter() {
        int xKanan;
        int xKiri;
        outloop:
        for (xKiri = 0; xKiri < lebarCitra; xKiri++) {
            for (int yKiri = 0; yKiri < tinggiCitra; yKiri++) {
                Color c = new Color(binaryImage.getRGB(xKiri, yKiri) & 0x00ffffff);
                if (c.getGreen() == 255) {
                    break outloop;
                }
            }
        }


        outloop2:
        for (xKanan = lebarCitra - 1; xKanan > 0; xKanan--) {
            for (int yKanan = 0; yKanan < tinggiCitra; yKanan++) {
                Color d = new Color(binaryImage.getRGB(xKanan, yKanan) & 0x00ffffff);
                if (d.getGreen() == 255) {
                    break outloop2;
                }
            }
        }

        this.diameterObjek = (xKanan - xKiri);
    }
}
