namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;

public class GeoPerdasTools
{
    public GeoPerdasTools()
    {
    }

    public static string mthdWriteLanguageDSSCommandCircuit(string strName, double dblBase_kV, double dblOperation_pu, string strBus, int intConstraintOperation)
    {
        return "New \"Circuit." + strName + "\" basekv=" + dblBase_kV.ToString().Replace(",", ".") + " pu=" + (intConstraintOperation == 1 ? mthdConstraintMaxOperation(dblOperation_pu).ToString().Replace(",", ".") : dblOperation_pu.ToString().Replace(",", ".")) + " bus1=\"" + strBus + "\" r1=0 x1=0.0001";
    }

    public static double mthdConstraintMaxOperation(double dblOperation_pu)
    {
        return dblOperation_pu > 1.05 ? 1.05 : dblOperation_pu;
    }

    public static string mthdWriteLanguageDSSCommandLoadshape(string strName, double PotAtvNorm01, double PotAtvNorm02, double PotAtvNorm03, double PotAtvNorm04, double PotAtvNorm05, double PotAtvNorm06, double PotAtvNorm07, double PotAtvNorm08, double PotAtvNorm09, double PotAtvNorm10, double PotAtvNorm11, double PotAtvNorm12, double PotAtvNorm13, double PotAtvNorm14, double PotAtvNorm15, double PotAtvNorm16, double PotAtvNorm17, double PotAtvNorm18, double PotAtvNorm19, double PotAtvNorm20, double PotAtvNorm21, double PotAtvNorm22, double PotAtvNorm23, double PotAtvNorm24)
    {
        return "New \"Loadshape." + strName + "\" 24 1.0 mult=(" + PotAtvNorm01.ToString().Replace(",", ".") + " " + PotAtvNorm02.ToString().Replace(",", ".") + " " + PotAtvNorm03.ToString().Replace(",", ".") + " " + PotAtvNorm04.ToString().Replace(",", ".") + " " + PotAtvNorm05.ToString().Replace(",", ".") + " " + PotAtvNorm06.ToString().Replace(",", ".") + " " + PotAtvNorm07.ToString().Replace(",", ".") + " " + PotAtvNorm08.ToString().Replace(",", ".") + " " + PotAtvNorm09.ToString().Replace(",", ".") + " " + PotAtvNorm10.ToString().Replace(",", ".") + " " + PotAtvNorm11.ToString().Replace(",", ".") + " " + PotAtvNorm12.ToString().Replace(",", ".") + " " + PotAtvNorm13.ToString().Replace(",", ".") + " " + PotAtvNorm14.ToString().Replace(",", ".") + " " + PotAtvNorm15.ToString().Replace(",", ".") + " " + PotAtvNorm16.ToString().Replace(",", ".") + " " + PotAtvNorm17.ToString().Replace(",", ".") + " " + PotAtvNorm18.ToString().Replace(",", ".") + " " + PotAtvNorm19.ToString().Replace(",", ".") + " " + PotAtvNorm20.ToString().Replace(",", ".") + " " + PotAtvNorm21.ToString().Replace(",", ".") + " " + PotAtvNorm22.ToString().Replace(",", ".") + " " + PotAtvNorm23.ToString().Replace(",", ".") + " " + PotAtvNorm24.ToString().Replace(",", ".") + ")";
    }

    public static string mthdWriteLanguageDSSCommandMTSwitch(string strName, string strCodFas, string strBus1, string strBus2)
    {
        return "New \"Line." + strName + "\" phases=" + mthdNumberPhasesSegments(strCodFas) + " bus1=\"" + strBus1 + mthdPhasing(strCodFas) + "\" bus2=\"" + strBus2 + mthdPhasing(strCodFas) + "\" r1=0.001 r0=0.001 x1=0 x0=0 c1=0 c0=0 length=0.001 switch=T";
    }

    public static string mthdWriteLanguageDSSCommandLine(string strName, string strCodFas, string strBus1, string strBus2, string strCodCond, double dblComp_km, string strProperty, int intConstraintLength, int intConstraintProperty)
    {
        if (strProperty == "TC" && intConstraintProperty == 1)
            return "New \"Line." + strName + "\" phases=" + mthdNumberPhasesSegments(strCodFas) + " bus1=\"" + strBus1 + mthdPhasing(strCodFas) + "\" bus2=\"" + strBus2 + mthdPhasing(strCodFas) + "\" r1=0.001 r0=0.001 x1=0 x0=0 c1=0 c0=0 length=" + (intConstraintLength == 1 ? mthdConstraintMaxLengthBranch(dblComp_km).ToString().Replace(",", ".") : dblComp_km.ToString().Replace(",", ".")) + " units=km switch=T";
        else
            return "New \"Line." + strName + "\" phases=" + mthdNumberPhasesSegments(strCodFas) + " bus1=\"" + strBus1 + mthdPhasing(strCodFas) + "\" bus2=\"" + strBus2 + mthdPhasing(strCodFas) + "\" linecode=\"" + strCodCond + "_" + mthdNumberPhasesSegments(strCodFas) + "\" length=" + (intConstraintLength == 1 ? mthdConstraintMaxLengthBranch(dblComp_km).ToString().Replace(",", ".") : dblComp_km.ToString().Replace(",", ".")) + " units=km";
    }

    public static string mthdWriteLanguageDSSCommandEnergymeter(string strName, string strElement)
    {
        return "New \"Energymeter." + strName + "\" element=\"" + strElement + "\" terminal=1";
    }

    /*  public static double mthdTransformerTotalPowerLoss_per(String strPhase, double dblPowerTransformer_kVA, double dblMaxVoltageTransformer_kV)
      {
          int intPhase = GeoPerdasTools.mthdNumberPhases(strPhase);
          double[,] dblTotalPowerLossDefault_per = new double[9, 6];

          // Monofásico (quando Bifásico basta multiplicar por 2) menor e igual que 15 kV (%)
          dblTotalPowerLossDefault_per[0, 0] = 2.80;
          dblTotalPowerLossDefault_per[1, 0] = 2.45;
          dblTotalPowerLossDefault_per[2, 0] = 2.20;
          dblTotalPowerLossDefault_per[3, 0] = 1.92;
          dblTotalPowerLossDefault_per[4, 0] = 1.77;
          dblTotalPowerLossDefault_per[5, 0] = 1.56;
          dblTotalPowerLossDefault_per[6, 0] = 1.48;
          dblTotalPowerLossDefault_per[7, 0] = 1.47;
          dblTotalPowerLossDefault_per[8, 0] = 1.45;

          // Monofásico (quando Bifásico basta multiplicar por 2) maior que 15 kV e menor ou igual que 24,2 kV (%)
          dblTotalPowerLossDefault_per[0, 1] = 3.10;
          dblTotalPowerLossDefault_per[1, 1] = 2.65;
          dblTotalPowerLossDefault_per[2, 1] = 2.43;
          dblTotalPowerLossDefault_per[3, 1] = 2.08;
          dblTotalPowerLossDefault_per[4, 1] = 1.97;
          dblTotalPowerLossDefault_per[5, 1] = 1.85;
          dblTotalPowerLossDefault_per[6, 1] = 1.61;
          dblTotalPowerLossDefault_per[7, 1] = 1.58;
          dblTotalPowerLossDefault_per[8, 1] = 1.50;

          // Monofásico (quando Bifásico basta multiplicar por 2) maior que 24,2 kV e menor ou igual que 36,2 kV (%)
          dblTotalPowerLossDefault_per[0, 2] = 3.20;
          dblTotalPowerLossDefault_per[1, 2] = 2.70;
          dblTotalPowerLossDefault_per[2, 2] = 2.50;
          dblTotalPowerLossDefault_per[3, 2] = 2.18;
          dblTotalPowerLossDefault_per[4, 2] = 1.97;
          dblTotalPowerLossDefault_per[5, 2] = 1.87;
          dblTotalPowerLossDefault_per[6, 2] = 1.63;
          dblTotalPowerLossDefault_per[7, 2] = 1.58;
          dblTotalPowerLossDefault_per[8, 2] = 1.48;

          // Trifásico menor e igual que 15 kV (%)
          dblTotalPowerLossDefault_per[0, 3] = 2.73;
          dblTotalPowerLossDefault_per[1, 3] = 2.32;
          dblTotalPowerLossDefault_per[2, 3] = 2.10;
          dblTotalPowerLossDefault_per[3, 3] = 1.86;
          dblTotalPowerLossDefault_per[4, 3] = 1.68;
          dblTotalPowerLossDefault_per[5, 3] = 1.56;
          dblTotalPowerLossDefault_per[6, 3] = 1.53;
          dblTotalPowerLossDefault_per[7, 3] = 1.45;
          dblTotalPowerLossDefault_per[8, 3] = 1.35;

          // Trifásico maior que 15 kV e menor e igual que 24,2 kV (%)
          dblTotalPowerLossDefault_per[0, 4] = 3.13;
          dblTotalPowerLossDefault_per[1, 4] = 2.63;
          dblTotalPowerLossDefault_per[2, 4] = 2.34;
          dblTotalPowerLossDefault_per[3, 4] = 2.07;
          dblTotalPowerLossDefault_per[4, 4] = 1.85;
          dblTotalPowerLossDefault_per[5, 4] = 1.74;
          dblTotalPowerLossDefault_per[6, 4] = 1.70;
          dblTotalPowerLossDefault_per[7, 4] = 1.60;
          dblTotalPowerLossDefault_per[8, 4] = 1.47;

          // Trifásico maior que 24,2 kV e menor e igual que 36,2 kV (%)
          dblTotalPowerLossDefault_per[0, 5] = 3.07;
          dblTotalPowerLossDefault_per[1, 5] = 2.58;
          dblTotalPowerLossDefault_per[2, 5] = 2.39;
          dblTotalPowerLossDefault_per[3, 5] = 2.11;
          dblTotalPowerLossDefault_per[4, 5] = 1.83;
          dblTotalPowerLossDefault_per[5, 5] = 1.76;
          dblTotalPowerLossDefault_per[6, 5] = 1.71;
          dblTotalPowerLossDefault_per[7, 5] = 1.60;
          dblTotalPowerLossDefault_per[8, 5] = 1.48;

          if (dblMaxVoltageTransformer_kV <= 15 && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblTotalPowerLossDefault_per[0, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblTotalPowerLossDefault_per[1, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[2, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblTotalPowerLossDefault_per[3, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblTotalPowerLossDefault_per[4, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblTotalPowerLossDefault_per[5, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[6, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblTotalPowerLossDefault_per[7, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblTotalPowerLossDefault_per[8, 0] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0181 * Math.Pow(dblPowerTransformer_kVA, 2) + 15.2 * dblPowerTransformer_kVA + 92.407) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 15 && dblMaxVoltageTransformer_kV <= 24.2) && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblTotalPowerLossDefault_per[0, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblTotalPowerLossDefault_per[1, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[2, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblTotalPowerLossDefault_per[3, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblTotalPowerLossDefault_per[4, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblTotalPowerLossDefault_per[5, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[6, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblTotalPowerLossDefault_per[7, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblTotalPowerLossDefault_per[8, 1] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0549 * Math.Pow(dblPowerTransformer_kVA, 2) + 19.624 * dblPowerTransformer_kVA + 71.175) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2) && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblTotalPowerLossDefault_per[0, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblTotalPowerLossDefault_per[1, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[2, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblTotalPowerLossDefault_per[3, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblTotalPowerLossDefault_per[4, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblTotalPowerLossDefault_per[5, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[6, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblTotalPowerLossDefault_per[7, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblTotalPowerLossDefault_per[8, 2] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0618 * Math.Pow(dblPowerTransformer_kVA, 2) + 20.167 * dblPowerTransformer_kVA + 75.016) / (1000 * dblPowerTransformer_kVA);
          }
          else if (dblMaxVoltageTransformer_kV <= 15 && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[0, 3];
              else if (dblPowerTransformer_kVA == 30)
                  return dblTotalPowerLossDefault_per[1, 3];
              else if (dblPowerTransformer_kVA == 45)
                  return dblTotalPowerLossDefault_per[2, 3];
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[3, 3];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblTotalPowerLossDefault_per[4, 3];
              else if (dblPowerTransformer_kVA == 150)
                  return dblTotalPowerLossDefault_per[5, 3];
              else if (dblPowerTransformer_kVA == 175)
                  return dblTotalPowerLossDefault_per[6, 3];
              else if (dblPowerTransformer_kVA == 225)
                  return dblTotalPowerLossDefault_per[7, 3];
              else if (dblPowerTransformer_kVA == 300)
                  return dblTotalPowerLossDefault_per[8, 3];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0107 * Math.Pow(dblPowerTransformer_kVA, 2) + 15.961 * dblPowerTransformer_kVA + 219.2) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 15 && dblMaxVoltageTransformer_kV <= 24.2) && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[0, 4];
              else if (dblPowerTransformer_kVA == 30)
                  return dblTotalPowerLossDefault_per[1, 4];
              else if (dblPowerTransformer_kVA == 45)
                  return dblTotalPowerLossDefault_per[2, 4];
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[3, 4];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblTotalPowerLossDefault_per[4, 4];
              else if (dblPowerTransformer_kVA == 150)
                  return dblTotalPowerLossDefault_per[5, 4];
              else if (dblPowerTransformer_kVA == 175)
                  return dblTotalPowerLossDefault_per[6, 4];
              else if (dblPowerTransformer_kVA == 225)
                  return dblTotalPowerLossDefault_per[7, 4];
              else if (dblPowerTransformer_kVA == 300)
                  return dblTotalPowerLossDefault_per[8, 4];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0141 * Math.Pow(dblPowerTransformer_kVA, 2) + 18.064 * dblPowerTransformer_kVA + 244.14) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2) && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblTotalPowerLossDefault_per[0, 5];
              else if (dblPowerTransformer_kVA == 30)
                  return dblTotalPowerLossDefault_per[1, 5];
              else if (dblPowerTransformer_kVA == 45)
                  return dblTotalPowerLossDefault_per[2, 5];
              else if (dblPowerTransformer_kVA == 75)
                  return dblTotalPowerLossDefault_per[3, 5];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblTotalPowerLossDefault_per[4, 5];
              else if (dblPowerTransformer_kVA == 150)
                  return dblTotalPowerLossDefault_per[5, 5];
              else if (dblPowerTransformer_kVA == 175)
                  return dblTotalPowerLossDefault_per[6, 5];
              else if (dblPowerTransformer_kVA == 225)
                  return dblTotalPowerLossDefault_per[7, 5];
              else if (dblPowerTransformer_kVA == 300)
                  return dblTotalPowerLossDefault_per[8, 5];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0134 * Math.Pow(dblPowerTransformer_kVA, 2) + 17.99 * dblPowerTransformer_kVA + 246.36) / (1000 * dblPowerTransformer_kVA);
          }

          return 0;
      }

      public static double mthdTransformerNoLoadPowerLoss_per(String strPhase, double dblPowerTransformer_kVA, double dblMaxVoltageTransformer_kV)
      {
          int intPhase = GeoPerdasTools.mthdNumberPhases(strPhase);
          double[,] dblNoLoadPowerLossDefault_per = new double[9, 6];

          // Monofásico (quando Bifásico basta multiplicar por 2) menor e igual que 15 kV (%)
          dblNoLoadPowerLossDefault_per[0, 0] = 0.70;
          dblNoLoadPowerLossDefault_per[1, 0] = 0.50;
          dblNoLoadPowerLossDefault_per[2, 0] = 0.43;
          dblNoLoadPowerLossDefault_per[3, 0] = 0.36;
          dblNoLoadPowerLossDefault_per[4, 0] = 0.36;
          dblNoLoadPowerLossDefault_per[5, 0] = 0.33;
          dblNoLoadPowerLossDefault_per[6, 0] = 0.27;
          dblNoLoadPowerLossDefault_per[7, 0] = 0.27;
          dblNoLoadPowerLossDefault_per[8, 0] = 0.26;

          // Monofásico (quando Bifásico basta multiplicar por 2) maior que 15 kV e menor ou igual que 24,2 kV (%)
          dblNoLoadPowerLossDefault_per[0, 1] = 0.80;
          dblNoLoadPowerLossDefault_per[1, 1] = 0.55;
          dblNoLoadPowerLossDefault_per[2, 1] = 0.50;
          dblNoLoadPowerLossDefault_per[3, 1] = 0.40;
          dblNoLoadPowerLossDefault_per[4, 1] = 0.39;
          dblNoLoadPowerLossDefault_per[5, 1] = 0.38;
          dblNoLoadPowerLossDefault_per[6, 1] = 0.30;
          dblNoLoadPowerLossDefault_per[7, 1] = 0.30;
          dblNoLoadPowerLossDefault_per[8, 1] = 0.28;

          // Monofásico (quando Bifásico basta multiplicar por 2) maior que 24,2 kV e menor ou igual que 36,2 kV (%)
          dblNoLoadPowerLossDefault_per[0, 2] = 0.90;
          dblNoLoadPowerLossDefault_per[1, 2] = 0.60;
          dblNoLoadPowerLossDefault_per[2, 2] = 0.53;
          dblNoLoadPowerLossDefault_per[3, 2] = 0.42;
          dblNoLoadPowerLossDefault_per[4, 2] = 0.40;
          dblNoLoadPowerLossDefault_per[5, 2] = 0.40;
          dblNoLoadPowerLossDefault_per[6, 2] = 0.32;
          dblNoLoadPowerLossDefault_per[7, 2] = 0.31;
          dblNoLoadPowerLossDefault_per[8, 2] = 0.28;

          // Trifásico menor e igual que 15 kV (%)
          dblNoLoadPowerLossDefault_per[0, 3] = 0.57;
          dblNoLoadPowerLossDefault_per[1, 3] = 0.50;
          dblNoLoadPowerLossDefault_per[2, 3] = 0.43;
          dblNoLoadPowerLossDefault_per[3, 3] = 0.39;
          dblNoLoadPowerLossDefault_per[4, 3] = 0.35;
          dblNoLoadPowerLossDefault_per[5, 3] = 0.32;
          dblNoLoadPowerLossDefault_per[6, 3] = 0.31;
          dblNoLoadPowerLossDefault_per[7, 3] = 0.29;
          dblNoLoadPowerLossDefault_per[8, 3] = 0.27;

          // Trifásico maior que 15 kV e menor e igual que 24,2 kV (%)
          dblNoLoadPowerLossDefault_per[0, 4] = 0.63;
          dblNoLoadPowerLossDefault_per[1, 4] = 0.53;
          dblNoLoadPowerLossDefault_per[2, 4] = 0.48;
          dblNoLoadPowerLossDefault_per[3, 4] = 0.42;
          dblNoLoadPowerLossDefault_per[4, 4] = 0.38;
          dblNoLoadPowerLossDefault_per[5, 4] = 0.35;
          dblNoLoadPowerLossDefault_per[6, 4] = 0.34;
          dblNoLoadPowerLossDefault_per[7, 4] = 0.32;
          dblNoLoadPowerLossDefault_per[8, 4] = 0.28;

          // Trifásico maior que 24,2 kV e menor e igual que 36,2 kV (%)
          dblNoLoadPowerLossDefault_per[0, 5] = 0.67;
          dblNoLoadPowerLossDefault_per[1, 5] = 0.55;
          dblNoLoadPowerLossDefault_per[2, 5] = 0.51;
          dblNoLoadPowerLossDefault_per[3, 5] = 0.43;
          dblNoLoadPowerLossDefault_per[4, 5] = 0.39;
          dblNoLoadPowerLossDefault_per[5, 5] = 0.36;
          dblNoLoadPowerLossDefault_per[6, 5] = 0.35;
          dblNoLoadPowerLossDefault_per[7, 5] = 0.33;
          dblNoLoadPowerLossDefault_per[8, 5] = 0.30;

          if (dblMaxVoltageTransformer_kV <= 15 && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblNoLoadPowerLossDefault_per[0, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblNoLoadPowerLossDefault_per[1, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[2, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblNoLoadPowerLossDefault_per[3, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblNoLoadPowerLossDefault_per[4, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblNoLoadPowerLossDefault_per[5, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[6, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblNoLoadPowerLossDefault_per[7, 0] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblNoLoadPowerLossDefault_per[8, 0] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0101 * Math.Pow(dblPowerTransformer_kVA, 2) + 3.3605 * dblPowerTransformer_kVA + 17.604) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 15 && dblMaxVoltageTransformer_kV <= 24.2) && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblNoLoadPowerLossDefault_per[0, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblNoLoadPowerLossDefault_per[1, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[2, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblNoLoadPowerLossDefault_per[3, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblNoLoadPowerLossDefault_per[4, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblNoLoadPowerLossDefault_per[5, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[6, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblNoLoadPowerLossDefault_per[7, 1] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblNoLoadPowerLossDefault_per[8, 1] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0132 * Math.Pow(dblPowerTransformer_kVA, 2) + 3.8514 * dblPowerTransformer_kVA + 19.033) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2) && (intPhase == 1 || intPhase == 2))
          {
              if (dblPowerTransformer_kVA == 5)
                  return dblNoLoadPowerLossDefault_per[0, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 10)
                  return dblNoLoadPowerLossDefault_per[1, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[2, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 25)
                  return dblNoLoadPowerLossDefault_per[3, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 37.5)
                  return dblNoLoadPowerLossDefault_per[4, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 50)
                  return dblNoLoadPowerLossDefault_per[5, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[6, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 85)
                  return dblNoLoadPowerLossDefault_per[7, 2] * intPhase;
              else if (dblPowerTransformer_kVA == 100)
                  return dblNoLoadPowerLossDefault_per[8, 2] * intPhase;
              else if (dblPowerTransformer_kVA < 100)
                  return 100 * intPhase * (-0.0151 * Math.Pow(dblPowerTransformer_kVA, 2) + 4.1019 * dblPowerTransformer_kVA + 20.978) / (1000 * dblPowerTransformer_kVA);
          }
          else if (dblMaxVoltageTransformer_kV <= 15 && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[0, 3];
              else if (dblPowerTransformer_kVA == 30)
                  return dblNoLoadPowerLossDefault_per[1, 3];
              else if (dblPowerTransformer_kVA == 45)
                  return dblNoLoadPowerLossDefault_per[2, 3];
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[3, 3];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblNoLoadPowerLossDefault_per[4, 3];
              else if (dblPowerTransformer_kVA == 150)
                  return dblNoLoadPowerLossDefault_per[5, 3];
              else if (dblPowerTransformer_kVA == 175)
                  return dblNoLoadPowerLossDefault_per[6, 3];
              else if (dblPowerTransformer_kVA == 225)
                  return dblNoLoadPowerLossDefault_per[7, 3];
              else if (dblPowerTransformer_kVA == 300)
                  return dblNoLoadPowerLossDefault_per[8, 3];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0027 * Math.Pow(dblPowerTransformer_kVA, 2) + 3.3358 * dblPowerTransformer_kVA + 48.012) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 15 && dblMaxVoltageTransformer_kV <= 24.2) && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[0, 4];
              else if (dblPowerTransformer_kVA == 30)
                  return dblNoLoadPowerLossDefault_per[1, 4];
              else if (dblPowerTransformer_kVA == 45)
                  return dblNoLoadPowerLossDefault_per[2, 4];
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[3, 4];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblNoLoadPowerLossDefault_per[4, 4];
              else if (dblPowerTransformer_kVA == 150)
                  return dblNoLoadPowerLossDefault_per[5, 4];
              else if (dblPowerTransformer_kVA == 175)
                  return dblNoLoadPowerLossDefault_per[6, 4];
              else if (dblPowerTransformer_kVA == 225)
                  return dblNoLoadPowerLossDefault_per[7, 4];
              else if (dblPowerTransformer_kVA == 300)
                  return dblNoLoadPowerLossDefault_per[8, 4];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0037 * Math.Pow(dblPowerTransformer_kVA, 2) + 3.8051 * dblPowerTransformer_kVA + 45.715) / (1000 * dblPowerTransformer_kVA);
          }
          else if ((dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2) && intPhase == 3)
          {
              if (dblPowerTransformer_kVA == 15)
                  return dblNoLoadPowerLossDefault_per[0, 5];
              else if (dblPowerTransformer_kVA == 30)
                  return dblNoLoadPowerLossDefault_per[1, 5];
              else if (dblPowerTransformer_kVA == 45)
                  return dblNoLoadPowerLossDefault_per[2, 5];
              else if (dblPowerTransformer_kVA == 75)
                  return dblNoLoadPowerLossDefault_per[3, 5];
              else if (dblPowerTransformer_kVA == 112.5)
                  return dblNoLoadPowerLossDefault_per[4, 5];
              else if (dblPowerTransformer_kVA == 150)
                  return dblNoLoadPowerLossDefault_per[5, 5];
              else if (dblPowerTransformer_kVA == 175)
                  return dblNoLoadPowerLossDefault_per[6, 5];
              else if (dblPowerTransformer_kVA == 225)
                  return dblNoLoadPowerLossDefault_per[7, 5];
              else if (dblPowerTransformer_kVA == 300)
                  return dblNoLoadPowerLossDefault_per[8, 5];
              else if (dblPowerTransformer_kVA < 300)
                  return 100 * (-0.0033 * Math.Pow(dblPowerTransformer_kVA, 2) + 3.8054 * dblPowerTransformer_kVA + 52.686) / (1000 * dblPowerTransformer_kVA);
          }

          return 0;
      }*/

    public static double mthdTransformerTotalPowerLoss_per(
      string strPhase,
      double dblPowerTransformer_kVA,
      double dblMaxVoltageTransformer_kV)
    {
        int num = mthdNumberPhases(strPhase);
        double[,] numArray = new double[9, 6];
        numArray[0, 0] = 2.8;
        numArray[1, 0] = 2.45;
        numArray[2, 0] = 2.2;
        numArray[3, 0] = 1.92;
        numArray[4, 0] = 1.77;
        numArray[5, 0] = 1.56;
        numArray[6, 0] = 1.48;
        numArray[7, 0] = 1.47;
        numArray[8, 0] = 1.45;
        numArray[0, 1] = 3.1;
        numArray[1, 1] = 2.65;
        numArray[2, 1] = 2.43;
        numArray[3, 1] = 2.08;
        numArray[4, 1] = 1.97;
        numArray[5, 1] = 1.85;
        numArray[6, 1] = 1.61;
        numArray[7, 1] = 1.58;
        numArray[8, 1] = 1.5;
        numArray[0, 2] = 3.2;
        numArray[1, 2] = 2.7;
        numArray[2, 2] = 2.5;
        numArray[3, 2] = 2.18;
        numArray[4, 2] = 1.97;
        numArray[5, 2] = 1.87;
        numArray[6, 2] = 1.63;
        numArray[7, 2] = 1.58;
        numArray[8, 2] = 1.48;
        numArray[0, 3] = 2.73;
        numArray[1, 3] = 2.32;
        numArray[2, 3] = 2.1;
        numArray[3, 3] = 1.86;
        numArray[4, 3] = 1.68;
        numArray[5, 3] = 1.56;
        numArray[6, 3] = 1.53;
        numArray[7, 3] = 1.45;
        numArray[8, 3] = 1.35;
        numArray[0, 4] = 3.13;
        numArray[1, 4] = 2.63;
        numArray[2, 4] = 2.34;
        numArray[3, 4] = 2.07;
        numArray[4, 4] = 1.85;
        numArray[5, 4] = 1.74;
        numArray[6, 4] = 1.7;
        numArray[7, 4] = 1.6;
        numArray[8, 4] = 1.47;
        numArray[0, 5] = 3.07;
        numArray[1, 5] = 2.58;
        numArray[2, 5] = 2.39;
        numArray[3, 5] = 2.11;
        numArray[4, 5] = 1.83;
        numArray[5, 5] = 1.76;
        numArray[6, 5] = 1.71;
        numArray[7, 5] = 1.6;
        numArray[8, 5] = 1.48;
        if (dblMaxVoltageTransformer_kV <= 15.0 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 0] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 0] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 0] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 0] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 0] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 0] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 0] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 0] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 0] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0181 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 15.2 * dblPowerTransformer_kVA + 92.407) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 15.0 && dblMaxVoltageTransformer_kV <= 24.2 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 1] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 1] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 1] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 1] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 1] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 1] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 1] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 1] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 1] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0549 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 19.624 * dblPowerTransformer_kVA + 71.175) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 2] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 2] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 2] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 2] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 2] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 2] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 2] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 2] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 2] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0618 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 20.167 * dblPowerTransformer_kVA + 75.016) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV <= 15.0 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 3];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 3];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 3];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 3];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 3];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 3];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 3];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 3];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 3];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0107 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 15.961 * dblPowerTransformer_kVA + 219.2) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 15.0 && dblMaxVoltageTransformer_kV <= 24.2 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 4];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 4];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 4];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 4];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 4];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 4];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 4];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 4];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 4];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0141 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 18.064 * dblPowerTransformer_kVA + 244.14) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 5];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 5];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 5];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 5];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 5];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 5];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 5];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 5];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 5];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0134 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 17.99 * dblPowerTransformer_kVA + 246.36) / (1000.0 * dblPowerTransformer_kVA);
        }
        return 0.0;
    }

    public static double mthdTransformerNoLoadPowerLoss_per(
      string strPhase,
      double dblPowerTransformer_kVA,
      double dblMaxVoltageTransformer_kV)
    {
        int num = mthdNumberPhases(strPhase);
        double[,] numArray = new double[9, 6];
        numArray[0, 0] = 0.7;
        numArray[1, 0] = 0.5;
        numArray[2, 0] = 0.43;
        numArray[3, 0] = 0.36;
        numArray[4, 0] = 0.36;
        numArray[5, 0] = 0.33;
        numArray[6, 0] = 0.27;
        numArray[7, 0] = 0.27;
        numArray[8, 0] = 0.26;
        numArray[0, 1] = 0.8;
        numArray[1, 1] = 0.55;
        numArray[2, 1] = 0.5;
        numArray[3, 1] = 0.4;
        numArray[4, 1] = 0.39;
        numArray[5, 1] = 0.38;
        numArray[6, 1] = 0.3;
        numArray[7, 1] = 0.3;
        numArray[8, 1] = 0.28;
        numArray[0, 2] = 0.9;
        numArray[1, 2] = 0.6;
        numArray[2, 2] = 0.53;
        numArray[3, 2] = 0.42;
        numArray[4, 2] = 0.4;
        numArray[5, 2] = 0.4;
        numArray[6, 2] = 0.32;
        numArray[7, 2] = 0.31;
        numArray[8, 2] = 0.28;
        numArray[0, 3] = 0.57;
        numArray[1, 3] = 0.5;
        numArray[2, 3] = 0.43;
        numArray[3, 3] = 0.39;
        numArray[4, 3] = 0.35;
        numArray[5, 3] = 0.32;
        numArray[6, 3] = 0.31;
        numArray[7, 3] = 0.29;
        numArray[8, 3] = 0.27;
        numArray[0, 4] = 0.63;
        numArray[1, 4] = 0.53;
        numArray[2, 4] = 0.48;
        numArray[3, 4] = 0.42;
        numArray[4, 4] = 0.38;
        numArray[5, 4] = 0.35;
        numArray[6, 4] = 0.34;
        numArray[7, 4] = 0.32;
        numArray[8, 4] = 0.28;
        numArray[0, 5] = 0.67;
        numArray[1, 5] = 0.55;
        numArray[2, 5] = 0.51;
        numArray[3, 5] = 0.43;
        numArray[4, 5] = 0.39;
        numArray[5, 5] = 0.36;
        numArray[6, 5] = 0.35;
        numArray[7, 5] = 0.33;
        numArray[8, 5] = 0.3;
        if (dblMaxVoltageTransformer_kV <= 15.0 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 0] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 0] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 0] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 0] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 0] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 0] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 0] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 0] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 0] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0101 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 3.3605 * dblPowerTransformer_kVA + 17.604) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 15.0 && dblMaxVoltageTransformer_kV <= 24.2 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 1] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 1] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 1] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 1] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 1] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 1] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 1] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 1] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 1] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0132 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 3.8514 * dblPowerTransformer_kVA + 19.033) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2 && (num == 1 || num == 2))
        {
            if (dblPowerTransformer_kVA == 5.0)
                return numArray[0, 2] * num;
            if (dblPowerTransformer_kVA == 10.0)
                return numArray[1, 2] * num;
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[2, 2] * num;
            if (dblPowerTransformer_kVA == 25.0)
                return numArray[3, 2] * num;
            if (dblPowerTransformer_kVA == 37.5)
                return numArray[4, 2] * num;
            if (dblPowerTransformer_kVA == 50.0)
                return numArray[5, 2] * num;
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[6, 2] * num;
            if (dblPowerTransformer_kVA == 85.0)
                return numArray[7, 2] * num;
            if (dblPowerTransformer_kVA == 100.0)
                return numArray[8, 2] * num;
            if (dblPowerTransformer_kVA < 100.0)
                return 100 * num * (-0.0151 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 4.1019 * dblPowerTransformer_kVA + 20.978) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV <= 15.0 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 3];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 3];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 3];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 3];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 3];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 3];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 3];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 3];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 3];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0027 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 3.3358 * dblPowerTransformer_kVA + 48.012) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 15.0 && dblMaxVoltageTransformer_kV <= 24.2 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 4];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 4];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 4];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 4];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 4];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 4];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 4];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 4];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 4];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0037 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 3.8051 * dblPowerTransformer_kVA + 45.715) / (1000.0 * dblPowerTransformer_kVA);
        }
        else if (dblMaxVoltageTransformer_kV > 24.2 && dblMaxVoltageTransformer_kV <= 36.2 && num == 3)
        {
            if (dblPowerTransformer_kVA == 15.0)
                return numArray[0, 5];
            if (dblPowerTransformer_kVA == 30.0)
                return numArray[1, 5];
            if (dblPowerTransformer_kVA == 45.0)
                return numArray[2, 5];
            if (dblPowerTransformer_kVA == 75.0)
                return numArray[3, 5];
            if (dblPowerTransformer_kVA == 112.5)
                return numArray[4, 5];
            if (dblPowerTransformer_kVA == 150.0)
                return numArray[5, 5];
            if (dblPowerTransformer_kVA == 175.0)
                return numArray[6, 5];
            if (dblPowerTransformer_kVA == 225.0)
                return numArray[7, 5];
            if (dblPowerTransformer_kVA == 300.0)
                return numArray[8, 5];
            if (dblPowerTransformer_kVA < 300.0)
                return 100.0 * (-0.0033 * Math.Pow(dblPowerTransformer_kVA, 2.0) + 3.8054 * dblPowerTransformer_kVA + 52.686) / (1000.0 * dblPowerTransformer_kVA);
        }
        return 0.0;
    }

    public static string mthdPhasing(string strPhases)
    {
        if (strPhases == "A")
            return ".1";
        else if (strPhases == "B")
            return ".2";
        else if (strPhases == "C")
            return ".3";
        else if (strPhases == "AN")
            return ".1.0";
        else if (strPhases == "BN")
            return ".2.0";
        else if (strPhases == "CN")
            return ".3.0";
        else if (strPhases == "AB")
            return ".1.2";
        else if (strPhases == "BC")
            return ".2.3";
        else if (strPhases == "CA")
            return ".3.1";
        else if (strPhases == "ABN")
            return ".1.2.0";
        else if (strPhases == "BCN")
            return ".2.3.0";
        else if (strPhases == "CAN")
            return ".3.1.0";
        else if (strPhases == "ABC")
            return ".1.2.3";
        else if (strPhases == "ABCN")
            return ".1.2.3.0";
        return "";
    }

    public static string mthdPhasingTertiary(string strPhases)
    {
        if (strPhases == "AN")
            return ".0.1";
        else if (strPhases == "BN")
            return ".0.2";
        else if (strPhases == "CN")
            return ".0.3";
        return "";
    }

    public static int mthdNumberPhases(string strPhases)
    {
        if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN")
            return 1;
        else if (strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
            return 2;
        else if (strPhases == "ABC" || strPhases == "ABCN")
            return 3;
        return 0;
    }

    public static string mthdNumberPhasesSegments(string strPhases)
    {
        if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN")
            return "1";
        else if (strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
            return "2";
        else if (strPhases == "ABC" || strPhases == "ABCN")
            return "3";
        return "";
    }

    public static string mthdNumberPhasesTransformers(string strPhases)
    {
        if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN" || strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
            return "1";
        else if (strPhases == "ABC" || strPhases == "ABCN")
            return "3";
        return "";
    }

    public static string mthdNumberPhasesRegulators(int intTypeRegul, string strPhases)
    {
        if (intTypeRegul == 4)
        {
            if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN")
                return "1";
            else if (strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
                return "2";
            else if (strPhases == "ABC" || strPhases == "ABCN")
                return "3";
        }
        else
        {
            if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN" || strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
                return "1";
            else if (strPhases == "ABC" || strPhases == "ABCN")
                return "3";
        }
        return "";
    }

    public static string mthdNumberPhasesLoads(string strPhases)
    {
        if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN" || strPhases == "AB" || strPhases == "BC" || strPhases == "CA" || strPhases == "ABN" || strPhases == "BCN" || strPhases == "CAN")
            return "1";
        else if (strPhases == "ABC" || strPhases == "ABCN")
            return "3";
        return "";
    }

    public static double mthdTensionLoads(string strPhases, double dblTensaoLine_kV, double dblTensaoPhase_kV)
    {
        if (strPhases == "A" || strPhases == "B" || strPhases == "C" || strPhases == "AN" || strPhases == "BN" || strPhases == "CN")
            return dblTensaoPhase_kV;
        else
            return dblTensaoLine_kV;
    }

    public static string mthdNamingBank(int intCodBnk)
    {
        if (intCodBnk == 1)
            return "A";
        else if (intCodBnk == 2)
            return "B";
        else if (intCodBnk == 3)
            return "C";
        else if (intCodBnk == 4)
            return "D";
        else if (intCodBnk == 5)
            return "E";
        else if (intCodBnk == 6)
            return "F";
        return "";
    }

    public static double mthdConstraintMaxLengthBranch(double dblLengthBranch)
    {
        return dblLengthBranch > 30 ? 30 : dblLengthBranch;
    }

    public static int mthdGetNumberDaysDUSADOMonth(int intMonth, int intYear, string strType)
    {
        int result = 0;
        DateTime dtCountDate;

        dtCountDate = new DateTime(intYear, intMonth, 1);
        while (dtCountDate.Year == intYear && dtCountDate.Month == intMonth)
        {
            if (strType == "DO" && (dtCountDate.DayOfWeek == DayOfWeek.Sunday || mthdIsHoliday(dtCountDate)))
                result++;
            else if (strType == "SA" && dtCountDate.DayOfWeek == DayOfWeek.Saturday && !mthdIsHoliday(dtCountDate))
                result++;
            else if (strType == "DU" && dtCountDate.DayOfWeek != DayOfWeek.Saturday && dtCountDate.DayOfWeek != DayOfWeek.Sunday && !mthdIsHoliday(dtCountDate))
                result++;
            dtCountDate = dtCountDate.AddDays(1);
        }

        return result;
    }

    public static string mthdGetTypeDayMonth(int intMonth, int intTypeDay)
    {
        return mthdGetTypeDay(intTypeDay) + string.Format("{0:00}", intMonth);
    }

    public static string mthdGetTypeDay(int intTypeDay)
    {
        if (intTypeDay == 1)
            return "DU";
        else if (intTypeDay == 2)
            return "SA";
        else if (intTypeDay == 3)
            return "DO";
        return "";
    }

    public static string mthdMonthToString(int intMonth)
    {
        if (intMonth == 1)
            return "JAN";
        else if (intMonth == 2)
            return "FEV";
        else if (intMonth == 3)
            return "MAR";
        else if (intMonth == 4)
            return "ABR";
        else if (intMonth == 5)
            return "MAI";
        else if (intMonth == 6)
            return "JUN";
        else if (intMonth == 7)
            return "JUL";
        else if (intMonth == 8)
            return "AGO";
        else if (intMonth == 9)
            return "SET";
        else if (intMonth == 10)
            return "OUT";
        else if (intMonth == 11)
            return "NOV";
        else if (intMonth == 12)
            return "DEZ";

        return "";
    }

    private static bool mthdIsHoliday(DateTime dtData)
    {
        int intDay, intMonth, intYear;

        intDay = dtData.Day;
        intMonth = dtData.Month;
        intYear = dtData.Year;
        if (intDay == 1 && intMonth == 1)
            return true;
        else if (intDay == 21 && intMonth == 4)
            return true;
        else if (intDay == 1 && intMonth == 5)
            return true;
        else if (intDay == 7 && intMonth == 9)
            return true;
        else if (intDay == 12 && intMonth == 10)
            return true;
        else if (intDay == 2 && intMonth == 11)
            return true;
        else if (intDay == 15 && intMonth == 11)
            return true;
        else if (intDay == 25 && intMonth == 12)
            return true;
        else if (mthdIsMovelHoliday(intDay, intMonth, intYear))
            return true;

        return false;
    }

    private static bool mthdIsMovelHoliday(int intDay, int intMonth, int intYear)
    {
        int temp01, temp02, temp03, temp04, temp05, temp06, temp07, temp08, temp09, temp10, temp11, temp12, temp13, temp14;
        DateTime dtInicialDate, dtHolidayDate1, dtHolidayDate2, dtHolidayDate3;

        temp01 = intYear % 19;
        temp02 = Convert.ToInt32(intYear / 100);
        temp03 = intYear % 100;
        temp04 = Convert.ToInt32(temp02 / 4);
        temp05 = temp02 % 4;
        temp06 = Convert.ToInt32((temp02 + 8) / 25);
        temp07 = Convert.ToInt32((temp02 - temp06 + 1) / 3);
        temp08 = (19 * temp01 + temp02 - temp04 - temp07 + 15) % 30;
        temp09 = Convert.ToInt32(temp03 / 4);
        temp10 = temp03 % 4;
        temp11 = (32 + 2 * temp05 + 2 * temp09 - temp08 - temp10) % 7;
        temp12 = Convert.ToInt32((temp01 + 11 * temp08 + 22 * temp11) / 451);
        temp13 = Convert.ToInt32((temp08 + temp11 - 7 * temp12 + 114) / 31);
        temp14 = (temp08 + temp11 - 7 * temp12 + 114) % 31;
        dtInicialDate = new DateTime(intYear, temp13, temp14 + 1);
        dtHolidayDate1 = dtInicialDate.AddDays(-47);
        dtHolidayDate2 = dtInicialDate.AddDays(-2);
        dtHolidayDate3 = dtInicialDate.AddDays(60);
        if (intDay == dtHolidayDate1.Day && intMonth == dtHolidayDate1.Month && intYear == dtHolidayDate1.Year)
            return true;
        if (intDay == dtHolidayDate2.Day && intMonth == dtHolidayDate2.Month && intYear == dtHolidayDate2.Year)
            return true;
        if (intDay == dtHolidayDate3.Day && intMonth == dtHolidayDate3.Month && intYear == dtHolidayDate3.Year)
            return true;

        return false;
    }

    public static string mthdLeftString(string strSource, int intLenght)
    {
        return strSource.Substring(0, Math.Min(intLenght, strSource.Length));
    }

    public static string mthdToStringUnicode(string strSource)
    {
        string[,] strTemp01 = { {"À", "&#192;"},
                                  {"Á", "&#193;"},
                                  {"Â", "&#194;"},
                                  {"Ã", "&#195;"},
                                  {"Ä", "&#196;"},
                                  {"Å", "&#197;"},
                                  {"Æ", "&#198;"},
                                  {"Ç", "&#199;"},
                                  {"È", "&#200;"},
                                  {"É", "&#201;"},
                                  {"Ê", "&#202;"},
                                  {"Ë", "&#203;"},
                                  {"Ì", "&#204;"},
                                  {"Í", "&#205;"},
                                  {"Î", "&#206;"},
                                  {"Ï", "&#207;"},
                                  {"Ð", "&#208;"},
                                  {"Ñ", "&#209;"},
                                  {"Ò", "&#210;"},
                                  {"Ó", "&#211;"},
                                  {"Ô", "&#212;"},
                                  {"Õ", "&#213;"},
                                  {"Ö", "&#214;"},
                                  {"×", "&#215;"},
                                  {"Ø", "&#216;"},
                                  {"Ù", "&#217;"},
                                  {"Ú", "&#218;"},
                                  {"Û", "&#219;"},
                                  {"Ü", "&#220;"},
                                  {"Ý", "&#221;"},
                                  {"Þ", "&#222;"},
                                  {"ß", "&#223;"},
                                  {"à", "&#224;"},
                                  {"á", "&#225;"},
                                  {"â", "&#226;"},
                                  {"ã", "&#227;"},
                                  {"ä", "&#228;"},
                                  {"å", "&#229;"},
                                  {"æ", "&#230;"},
                                  {"ç", "&#231;"},
                                  {"è", "&#232;"},
                                  {"é", "&#233;"},
                                  {"ê", "&#234;"},
                                  {"ë", "&#235;"},
                                  {"ì", "&#236;"},
                                  {"í", "&#237;"},
                                  {"î", "&#238;"},
                                  {"ï", "&#239;"},
                                  {"ð", "&#240;"},
                                  {"ñ", "&#241;"},
                                  {"ò", "&#242;"},
                                  {"ó", "&#243;"},
                                  {"ô", "&#244;"},
                                  {"õ", "&#245;"},
                                  {"ö", "&#246;"},
                                  {"÷", "&#247;"},
                                  {"ø", "&#248;"},
                                  {"ù", "&#249;"},
                                  {"ú", "&#250;"},
                                  {"û", "&#251;"},
                                  {"ü", "&#252;"},
                                  {"ý", "&#253;"},
                                  {"þ", "&#254;"},
                                  {"ÿ", "&#255;"},
                                  {"Ā", "&#256;"},
                                  {"ā", "&#257;"},
                                  {"Ă", "&#258;"},
                                  {"ă", "&#259;"},
                                  {"Ą", "&#260;"},
                                  {"ą", "&#261;"},
                                  {"Ć", "&#262;"},
                                  {"ć", "&#263;"},
                                  {"Ĉ", "&#264;"},
                                  {"ĉ", "&#265;"},
                                  {"Ċ", "&#266;"},
                                  {"ċ", "&#267;"},
                                  {"Č", "&#268;"},
                                  {"č", "&#269;"},
                                  {"Ď", "&#270;"},
                                  {"ď", "&#271;"},
                                  {"Đ", "&#272;"},
                                  {"đ", "&#273;"},
                                  {"Ē", "&#274;"},
                                  {"ē", "&#275;"},
                                  {"Ĕ", "&#276;"},
                                  {"ĕ", "&#277;"},
                                  {"Ė", "&#278;"},
                                  {"ė", "&#279;"},
                                  {"Ę", "&#280;"},
                                  {"ę", "&#281;"},
                                  {"Ě", "&#282;"},
                                  {"ě", "&#283;"},
                                  {"Ĝ", "&#284;"},
                                  {"ĝ", "&#285;"},
                                  {"Ğ", "&#286;"},
                                  {"ğ", "&#287;"},
                                  {"Ġ", "&#288;"},
                                  {"ġ", "&#289;"},
                                  {"Ģ", "&#290;"},
                                  {"ģ", "&#291;"},
                                  {"Ĥ", "&#292;"},
                                  {"ĥ", "&#293;"},
                                  {"Ħ", "&#294;"},
                                  {"ħ", "&#295;"},
                                  {"Ĩ", "&#296;"},
                                  {"ĩ", "&#297;"},
                                  {"Ī", "&#298;"},
                                  {"ī", "&#299;"},
                                  {"Ĭ", "&#300;"},
                                  {"ĭ", "&#301;"},
                                  {"Į", "&#302;"},
                                  {"į", "&#303;"},
                                  {"İ", "&#304;"},
                                  {"ı", "&#305;"},
                                  {"Ĳ", "&#306;"},
                                  {"ĳ", "&#307;"},
                                  {"Ĵ", "&#308;"},
                                  {"ĵ", "&#309;"},
                                  {"Ķ", "&#310;"},
                                  {"ķ", "&#311;"},
                                  {"ĸ", "&#312;"},
                                  {"Ĺ", "&#313;"},
                                  {"ĺ", "&#314;"},
                                  {"Ļ", "&#315;"},
                                  {"ļ", "&#316;"},
                                  {"Ľ", "&#317;"},
                                  {"ľ", "&#318;"},
                                  {"Ŀ", "&#319;"},
                                  {"ŀ", "&#320;"},
                                  {"Ł", "&#321;"},
                                  {"ł", "&#322;"},
                                  {"Ń", "&#323;"},
                                  {"ń", "&#324;"},
                                  {"Ņ", "&#325;"},
                                  {"ņ", "&#326;"},
                                  {"Ň", "&#327;"},
                                  {"ň", "&#328;"},
                                  {"ŉ", "&#329;"},
                                  {"Ŋ", "&#330;"},
                                  {"ŋ", "&#331;"},
                                  {"Ō", "&#332;"},
                                  {"ō", "&#333;"},
                                  {"Ŏ", "&#334;"},
                                  {"ŏ", "&#335;"} };

        for (int i = 0; i <= strTemp01.GetLength(0) - 1; i++)
            strSource = strSource.Replace(strTemp01[i, 0], strTemp01[i, 1]);

        return strSource;
    }

    public static double mthdNumberCorrection(double dblValue)
    {
        return double.IsInfinity(dblValue) || double.IsNaN(dblValue) ? 0 : Math.Abs(dblValue) < 0.000000001 ? 0 : dblValue > 99999999.99999999 ? 0 : dblValue < -99999999.99999999 ? 0 : dblValue;
    }

    public static double mthdNumberCorrectionBD(object objValue)
    {
        return DBNull.Value.Equals(objValue) ? 0 : Convert.ToDouble(objValue);
    }

    public static string mthdQuerySQLFeeders(string strFeeders)
    {
        string result = "";
        int i = 0;
        string[] strArrayFeeders = strFeeders.Split(';');

        foreach (string strFeeder in strArrayFeeders)
        {
            if (i == 0)
                result += "(t1.CodAlim = '" + strFeeder + "'";
            else
                result += " OR t1.CodAlim = '" + strFeeder + "'";
            i++;
        }
        result += ")";

        return result;
    }
}