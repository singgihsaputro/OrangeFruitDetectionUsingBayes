/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package tkbayes;

/**
 *
 * @author sion
 */
public class Datatraining {

    final int IMG_WIDTH = 427, IMG_HEIGHT = 367;
    final int jumlah = 15;
    double rata2Lemon[] = new double[4], rata2Manis[] = new double[4], rata2Nipis[] = new double[4];
    double invers_covarian[][] = new double[4][4];
    double matriksLemon[][];
    double matriksManis[][];
    double matriksNipis[][];
    double rt2red, rt2green, rt2blue, rtd;

    public void covarianGlobal(double[][] a, double[][] b, double[][] c) {
        matriksLemon = a;
        matriksManis = b;
        matriksNipis = c;
        rt2red = 0;
        rt2green = 0;
        rt2blue = 0;
        rtd = 0;
        for (int i = 0; i < matriksNipis.length; i++) {
            rt2red += matriksNipis[i][0];
            rt2green += matriksNipis[i][1];
            rt2blue += matriksNipis[i][2];
            rtd += matriksNipis[i][3];
        }
        for (int i = 0; i < matriksManis.length; i++) {
            rt2red += matriksManis[i][0];
            rt2green += matriksManis[i][1];
            rt2blue += matriksManis[i][2];
            rtd += matriksManis[i][3];
        }
        for (int i = 0; i < matriksLemon.length; i++) {
            rt2red += matriksLemon[i][0];
            rt2green += matriksLemon[i][1];
            rt2blue += matriksLemon[i][2];
            rtd += matriksLemon[i][3];
        }
//menghitung mean global

        rt2red = rt2red / jumlah;
        rt2green = rt2green / jumlah;
        rt2blue = rt2blue / jumlah;
        rtd = rtd / jumlah;

        //mean kelas lemon
        for (int i = 0; i < 3; i++) {
            rata2Lemon[0] += matriksLemon[i][0];
            rata2Lemon[1] += matriksLemon[i][1];
            rata2Lemon[2] += matriksLemon[i][2];
            rata2Lemon[3] += matriksLemon[i][3];
            // menghitung zero mean
            matriksLemon[i][0] -= rt2red;
            matriksLemon[i][1] -= rt2green;
            matriksLemon[i][2] -= rt2blue;
            matriksLemon[i][3] -= rtd;
        }


        rata2Lemon[0] = rata2Lemon[0] / 3;
        rata2Lemon[1] = rata2Lemon[1] / 3;
        rata2Lemon[2] = rata2Lemon[2] / 3;
        rata2Lemon[3] = rata2Lemon[3] / 3;

        //mean kelas manis
        for (int i = 0; i < 5; i++) {
            rata2Manis[0] += matriksManis[i][0];
            rata2Manis[1] += matriksManis[i][1];
            rata2Manis[2] += matriksManis[i][2];
            rata2Manis[3] += matriksManis[i][3];
            matriksManis[i][0] -= rt2red;
            matriksManis[i][1] -= rt2green;
            matriksManis[i][2] -= rt2blue;
            matriksManis[i][3] -= rtd;
        }

        rata2Manis[0] = rata2Manis[0] / 5;
        rata2Manis[1] = rata2Manis[1] / 5;
        rata2Manis[2] = rata2Manis[2] / 5;
        rata2Manis[3] = rata2Manis[3] / 5;

        //mean kelas nipis
        for (int i = 0; i < 7; i++) {
            rata2Nipis[0] += matriksNipis[i][0];
            rata2Nipis[1] += matriksNipis[i][1];
            rata2Nipis[2] += matriksNipis[i][2];
            rata2Nipis[3] += matriksNipis[i][3];
            matriksNipis[i][0] -= rt2red;
            matriksNipis[i][1] -= rt2green;
            matriksNipis[i][2] -= rt2blue;
            matriksNipis[i][3] -= rtd;
        }
        rata2Nipis[0] = rata2Nipis[0] / 7;
        rata2Nipis[1] = rata2Nipis[1] / 7;
        rata2Nipis[2] = rata2Nipis[2] / 7;
        rata2Nipis[3] = rata2Nipis[3] / 7;

        double kovarian[][] = new double[4][4];

        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 4; j++) {
                kovarian[i][j] = covarian(matriksLemon, 4, 3)[i][j] * 3 / jumlah
                        + covarian(matriksManis, 4, 5)[i][j] * 5 / jumlah
                        + covarian(matriksNipis, 4, 7)[i][j] * 7 / jumlah;
            }
        }
        invers_covarian = inverse(kovarian);
    }

    public double[][] transpose(double m[][], int x, int y) {
        double hasil[][] = new double[x][y];
        for (int i = 0; i < y; i++) {
            for (int j = 0; j < x; j++) {
                hasil[j][i] = m[i][j];
            }
        }
        return hasil;
    }

    public double[][] covarian(double m[][], int x, int y) {
        double hasil[][] = new double[x][x];
        double transpose[][] = transpose(m, x, y);
        for (int i = 0; i < x; i++) {
            for (int j = 0; j < x; j++) {
                for (int k = 0; k < y; k++) {
                    hasil[i][j] += transpose[i][k] * m[k][j];
                }
                hasil[i][j] = hasil[i][j] / y;
            }
        }
        return hasil;
    }

    public double determinant(double[][] matrix) {
        if (matrix.length == 1) {
            return matrix[0][0];
        }
        if (matrix.length == 2) {
            return (matrix[0][0] * matrix[1][1]) - (matrix[0][1] * matrix[1][0]);
        }
        double sum = 0.0;
        for (int i = 0; i < matrix.length; i++) {
            sum += changeSign(i) * matrix[0][i] * determinant(createSubMatrix(matrix, 0, i));
        }
        return sum;
    }

    public int changeSign(int a) {
        if (a % 2 == 0) {
            return 1;
        } else {
            return -1;
        }
    }

    public double[][] createSubMatrix(double[][] matrix, int excluding_row, int excluding_col) {
        double[][] mat = new double[matrix.length - 1][matrix.length - 1];
        int r = -1;
        for (int i = 0; i < matrix.length; i++) {
            if (i == excluding_row) {
                continue;
            }
            r++;
            int c = -1;
            for (int j = 0; j < matrix.length; j++) {
                if (j == excluding_col) {
                    continue;
                }
                mat[r][++c] = matrix[i][j];
            }
        }
        return mat;
    }

    public double[][] cofactor(double[][] matrix) {
        double mat[][] = new double[matrix.length][matrix.length];
        for (int i = 0; i < matrix.length; i++) {
            for (int j = 0; j < matrix.length; j++) {
                mat[i][j] = changeSign(i) * changeSign(j) * determinant(createSubMatrix(matrix, i, j));
            }
        }
        return mat;
    }

    public double[][] inverse(double[][] matrix) {
        double mat[][] = transpose(cofactor(matrix), matrix.length, matrix.length);
        for (int i = 0; i < mat.length; i++) {
            for (int j = 0; j < mat.length; j++) {
                mat[i][j] = mat[i][j] / determinant(matrix);
            }
        }
        return mat;
    }
}
