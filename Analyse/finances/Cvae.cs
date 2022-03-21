using System;

namespace finances
{
    public class Cvae
    {
        const double MIN_CVAE = 125;

        public static double ComputeRatio(double ca)
        {
            if (ca < 500000)
            {
                return 0;
            }
            if (ca < 3E6)
            {
                return 0.75 * (ca - 500E3) / 2.5E6;
            }
            if (ca < 10E6)
            {
                return 0.25 + 0.45 * (ca - 3E6) / 7E6;
            }
            if (ca < 50E6)
            {
                return 0.7 + 0.05 * (ca - 10E6) / 40E6;
            }
            return 0.75;
        }

        public static double ComputeTax(double ca, double va)
        {
            double taux = ComputeRatio(ca);
            double cvae = taux * va / 100;
            if (ca < 2E6)
            {
                cvae = cvae - 500;
            }
            if (cvae < 125)
            {
                cvae = 125;
            }

            double additionalTax = 0.0346 * cvae;
            double gestion = 0.01 * (cvae + additionalTax);
            //Console.WriteLine($"cvae = {cvae}, taxe additionnelle = {additionalTax}, frais de gestion = {gestion}");
            return cvae + additionalTax + gestion;
        }
    }
}


