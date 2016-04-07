/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package tkbayes;

/**
 *
 * @author sion
 */
public class datatesting {

    Datatraining pdt;
    double[] fitur = new double[4];
    double f1 = 0, f2 = 0, f3 = 0;
    double priorlemon = 0.2, priormanis = 0.33, priornipis = 0.47;
    double posteriorlemon, posteriormanis, posteriornipis;

    public String Pengenalan_Jeruk_Bayes(double a, double b, double c, double d, Datatraining e) {

        System.out.println("A" + a);
        System.out.println("B" + b);
        System.out.println("C" + c);
        System.out.println("D" + d);
        this.pdt = e;
        fitur[0] = a;
        fitur[1] = b;
        fitur[2] = c;
        fitur[3] = d;

        double temp[] = new double[4];
        double temp2[] = new double[4];
        double temp3[] = new double[4];
        double temp4[] = new double[4];
        double temp5[] = new double[4];
        double temp6[] = new double[4];
        double x = 0, x2 = 0, x3 = 0;

        for (int i = 0; i < 4; i++) {
            temp[i] = fitur[i] - e.rata2Lemon[i];
        }
        for (int j = 0; j < 4; j++) {
            for (int k = 0; k < 4; k++) {
                temp2[j] += temp[k] * e.inverse(e.covarian(e.matriksLemon, 4, 3))[k][j];
            }
        }
        for (int i = 0; i < 4; i++) {
            x += temp2[i] * temp[i];
        }
        Double detlem = e.determinant(e.covarian(e.matriksLemon, 4, 3));
        if (detlem < 0) {
            detlem = detlem * -1;
        }
        f1 = 2 * Math.pow(Math.PI, -2) * Math.pow(detlem, -0.5) * Math.pow(Math.E, -x / 2);
        System.out.println(f1);
        System.out.println("determinan" + detlem);
        System.out.println("epangkatlem : " + Math.pow(Math.E, -x / 2));
        System.out.println("like lemon " + f1);
        System.out.println("X Lemon=" + x);
        System.out.println("matrik covarian");
        double[][] aa = e.covarian(e.matriksLemon, 4, 3);
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 4; j++) {
                System.out.print(" " + aa[i][j]);
            }
            System.out.println();

        }
        System.out.println("");
        System.out.println("mean global red : " + e.rt2red);
        System.out.println("mean global green : " + e.rt2green);
        System.out.println("mean global blue : " + e.rt2blue);
        System.out.println("mean global diameter : " + e.rtd);
        System.out.println("");
        System.out.println("matrik zero mean lemon");
        double[][] aa2 = e.matriksLemon;
        for (int i = 0; i < 3; i++) {
            for (int j = 0; j < 4; j++) {
                System.out.print(" " + aa2[i][j]);
            }
            System.out.println();

        }

        System.out.println("");
        System.out.println("matrik transpose lemon");
        double[][] aa1 = e.transpose(e.matriksLemon, 4, 3);
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 3; j++) {
                System.out.print(" " + aa1[i][j]);
            }
            System.out.println();

        }

        for (int i = 0; i < 4; i++) {
            temp3[i] = fitur[i] - e.rata2Manis[i];
        }
        for (int j = 0; j < 4; j++) {
            for (int k = 0; k < 4; k++) {
                temp4[j] += temp[k] * e.inverse(e.covarian(e.matriksManis, 4, 5))[k][j];
            }
        }
        for (int i = 0; i < 4; i++) {
            x2 += temp4[i] * temp3[i];
        }
        Double detman = e.determinant(e.covarian(e.matriksManis, 4, 5));
        if (detman < 0) {
            detman = detman * -1;
        }
        f2 = 2 * Math.pow(Math.PI, -2) * Math.pow(detman, -0.5) * Math.pow(Math.E, -x2 / 2);
        System.out.println(f2);
        System.out.println("determinan " + detman);
        System.out.println("pangkat : " + Math.pow(Math.E, -x2 / 2));
        System.out.println("aaaa : " + Math.pow(e.determinant(e.covarian(e.matriksManis, 4, 5)), -0.5));

        System.out.println("X manis=" + x2);
        System.out.println("like manis " + f2);
        for (int i = 0; i < 4; i++) {
            temp5[i] = fitur[i] - e.rata2Nipis[i];
        }
        for (int j = 0; j < 4; j++) {
            for (int k = 0; k < 4; k++) {
                temp6[j] += temp[k] * e.inverse(e.covarian(e.matriksNipis, 4, 7))[k][j];
            }
        }
        for (int i = 0; i < 4; i++) {
            x3 += temp6[i] * temp5[i];
        }
        Double detnip = e.determinant(e.covarian(e.matriksNipis, 4, 7));
        if (detnip < 0) {
            detnip = detnip * -1;
        }
        System.out.println("determinan " + detnip);
        f3 = 2 * Math.pow(Math.PI, -2) * Math.pow(detnip, -0.5) * Math.pow(Math.E, -x3 / 2);

        System.out.println("X nipis=" + x3);
        System.out.println("like nipis " + f3);

        //menghitung posterior
        posteriorlemon = (f1 * priorlemon);
        posteriormanis = (f2 * priormanis);
        posteriornipis = (f3 * priornipis);
        if (posteriorlemon > posteriormanis && posteriorlemon > posteriornipis) {
            return "Jeruk Lemon";
        } else if (posteriormanis > posteriorlemon && posteriormanis > posteriornipis) {
            return "Jeruk Manis";
        } else {
            return "Jeruk Nipis";
        }
    }
}
