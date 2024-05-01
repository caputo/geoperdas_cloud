using System.Collections;

namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;

public class EnergyBalanceTools
{
    public EnergyBalanceTools()
    {
        setInjectEnergy_MWh();
        setSupplyEnergy_Mwh();
        setSupplyEnergywithoutNet_MWh();
        setSupplyEnergywithNet_MWh();
        setTecnicalLossesFixedEnergy_MWh();
        setTecnicalLossesTransformersEnergy_MWh();
        setTransformer();
        setSegmentTransformation();
        setProportionTtoST();
        setProportionSTtoT();
        setEnergyTransformerOutLevelDownUp_MWh();
        setEnergyTransformerEnterLevelDownUp_MWh();
        setEnergyTransformerEnterLevelUpDown_MWh();
        setEnergyTransformerOutLevelUpDown_MWh();
    }

    // Controle de ativação do nível
    private int[] inIsActiveLevel = new int[5];

    private void setIsActiveLevel()
    {
        for (int i = 0; i < inIsActiveLevel.GetLength(0); i++)
        {
            if (dblInjectEnergy_MWh[i] != 0 || dblSupplyEnergy_MWh[i] != 0 || dblSupplyEnergywithoutNet_MWh[i] != 0 || dblTecnicalLossesFixedEnergy_MWh[i] != 0)
            {
                inIsActiveLevel[i] = 1;
                break;
            }
            else
            {
                inIsActiveLevel[i] = 0;
            }
        }
    }

    // Energia injetada no nível
    private double[] dblInjectEnergy_MWh = new double[5];

    public void setInjectEnergy_MWh()
    {
        for (int i = 0; i < dblInjectEnergy_MWh.GetLength(0); i++)
        {
            dblInjectEnergy_MWh[i] = 0;
        }
    }

    public void addInjectEnergy_MWh(string strSegmet, double dblEnergy_MWh)
    {
        int i = getLevelNumber(strSegmet);

        if (i != -1 && dblEnergy_MWh != 0)
        {
            inIsActiveLevel[i] = 1;
            dblInjectEnergy_MWh[i] += dblEnergy_MWh;
        }
    }

    // Energia fornecida no nível
    private double[] dblSupplyEnergy_MWh = new double[5];

    public void setSupplyEnergy_Mwh()
    {
        for (int i = 0; i < dblSupplyEnergy_MWh.GetLength(0); i++)
        {
            dblSupplyEnergy_MWh[i] = 0;
        }
    }

    public void addSupplyEnergy_MWh(string strSegmet, double dblEnergy_MWh)
    {
        int i = getLevelNumber(strSegmet);

        if (i != -1 && dblEnergy_MWh != 0)
        {
            inIsActiveLevel[i] = 1;
            dblSupplyEnergy_MWh[i] += dblEnergy_MWh;
        }
    }

    // Energia fornecida sem rede associada
    private double[] dblSupplyEnergywithoutNet_MWh = new double[5];

    public void setSupplyEnergywithoutNet_MWh()
    {
        for (int i = 0; i < dblSupplyEnergywithoutNet_MWh.GetLength(0); i++)
        {
            dblSupplyEnergywithoutNet_MWh[i] = 0;
        }
    }

    public void addSupplyEnergywithoutNet_MWh(string strSegmet, double dblEnergy_MWh)
    {
        int i = getLevelNumber(strSegmet);

        if (i != -1 && dblEnergy_MWh != 0)
        {
            inIsActiveLevel[i] = 1;
            dblSupplyEnergywithoutNet_MWh[i] += dblEnergy_MWh;
        }
    }

    // Energia fornecida com rede associada
    private double[] dblSupplyEnergywithNet_MWh = new double[5];

    private void setSupplyEnergywithNet_MWh()
    {
        for (int i = 0; i < dblSupplyEnergywithNet_MWh.GetLength(0); i++)
        {
            dblSupplyEnergywithNet_MWh[i] = dblSupplyEnergy_MWh[i] - dblSupplyEnergywithoutNet_MWh[i];
        }
    }

    // Energia do nível no balanço
    private double[] dblBalanceLevelEnergy_MWh = new double[5];

    private void setBalanceLevelEnergy_MWh()
    {
        for (int i = 0; i < dblBalanceLevelEnergy_MWh.GetLength(0); i++)
        {
            dblBalanceLevelEnergy_MWh[i] = dblInjectEnergy_MWh[i] - dblSupplyEnergywithoutNet_MWh[i] - dblEnergyTransformerOutLevelDownUp_MWh[i] + dblEnergyTransformerEnterLevelDownUp_MWh[i] + dblEnergyTransformerEnterLevelUpDown_MWh[i];
        }
    }

    private void setBalanceLevelEnergy_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblBalanceLevelEnergy_MWh[intSegmet] = dblInjectEnergy_MWh[intSegmet] - dblSupplyEnergywithoutNet_MWh[intSegmet] - dblEnergyTransformerOutLevelDownUp_MWh[intSegmet] + dblEnergyTransformerEnterLevelDownUp_MWh[intSegmet] + dblEnergyTransformerEnterLevelUpDown_MWh[intSegmet];
        }
    }

    // Fator de correção do nível
    private double[] dblCorretionFatorLevel_pu = new double[5];

    // Perdas de energia do nível não corrigível
    private double[] dblTecnicalLossesFixedEnergy_MWh = new double[5];

    public void setTecnicalLossesFixedEnergy_MWh()
    {
        for (int i = 0; i < dblTecnicalLossesFixedEnergy_MWh.GetLength(0); i++)
        {
            dblTecnicalLossesFixedEnergy_MWh[i] = 0;
        }
    }

    public void addTecnicalLossesFixedEnergy_MWh(string strSegmet, double dblTecnicalLosses_MWh)
    {
        int i = getLevelNumber(strSegmet);

        if (i != -1 && dblTecnicalLosses_MWh != 0)
        {
            inIsActiveLevel[i] = 1;
            dblTecnicalLossesFixedEnergy_MWh[i] += dblTecnicalLosses_MWh;
        }
    }

    // Perdas de energia do nível relativa aos transformadores de potência
    private double[] dblTecnicalLossesTransformersEnergy_MWh = new double[5];

    private void setTecnicalLossesTransformersEnergy_MWh()
    {
        for (int i = 0; i < dblTecnicalLossesTransformersEnergy_MWh.GetLength(0); i++)
        {
            dblTecnicalLossesTransformersEnergy_MWh[i] = 0;
        }
    }

    private void setTecnicalLossesTransformersEnergy_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblTecnicalLossesTransformersEnergy_MWh[intSegmet] = 0;
            for (int i = 0; i < intPrimTransformer.Count; i++)
            {
                if ((int)intAllocationTransformer[i] == intSegmet)
                    dblTecnicalLossesTransformersEnergy_MWh[intSegmet] += (double)dblIronLossesTransformer_MWh[i] + (double)dblCopperLossesCorrectedTransformer_MWh[i];
            }
        }
    }

    private double[] dblTecnicalLossesVariableBeforeEnergy_MWh = new double[5];
    private double[] dblTecnicalLossesVariableAfterEnergy_MWh = new double[5];
    private double[] dblTecnicalLossesMeterEnergy_MWh = new double[5];

    // Diferença do balanço
    private double[] dblDiferenceBalanceEnergy_MWh = new double[5];

    private void setDiferenceBalanceEnergy_MWh()
    {
        for (int i = 0; i < dblDiferenceBalanceEnergy_MWh.GetLength(0); i++)
        {
            dblDiferenceBalanceEnergy_MWh[i] = dblBalanceLevelEnergy_MWh[i] - (dblTecnicalLossesFixedEnergy_MWh[i] - dblTecnicalLossesTransformersEnergy_MWh[i]) - dblEnergyTransformerOutLevelUpDown_MWh[i] - dblSupplyEnergywithNet_MWh[i];
        }
    }

    private void setDiferenceBalanceEnergy_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblDiferenceBalanceEnergy_MWh[intSegmet] = dblBalanceLevelEnergy_MWh[intSegmet] - (dblTecnicalLossesFixedEnergy_MWh[intSegmet] - dblTecnicalLossesTransformersEnergy_MWh[intSegmet]) - dblEnergyTransformerOutLevelUpDown_MWh[intSegmet] - dblSupplyEnergywithNet_MWh[intSegmet];
        }
    }

    // Transformadores de potência
    private ArrayList intPrimTransformer = new ArrayList();

    private ArrayList intSecuTransformer = new ArrayList();
    private ArrayList intTercTransformer = new ArrayList();
    private ArrayList dblEnergyTransformer_MWh = new ArrayList();
    private ArrayList dblIronLossesTransformer_MWh = new ArrayList();
    private ArrayList dblCopperLossesTransformer_MWh = new ArrayList();
    private ArrayList intAllocationTransformer = new ArrayList();
    private ArrayList dblEnergyTransformerBalanceBefore_MWh = new ArrayList();
    private ArrayList dblProportionTransformer_pu = new ArrayList();
    private ArrayList dblEnergyTransformerBalanceAfter_MWh = new ArrayList();
    private ArrayList dblCorretionFatorTransformer_pu = new ArrayList();
    private ArrayList dblCopperLossesCorrectedTransformer_MWh = new ArrayList();
    private ArrayList dblEnergyTransformerResult_MWh = new ArrayList();
    private ArrayList intFlow = new ArrayList();

    public void setTransformer()
    {
        intPrimTransformer.Clear();
        intSecuTransformer.Clear();
        intTercTransformer.Clear();
        dblEnergyTransformer_MWh.Clear();
        dblIronLossesTransformer_MWh.Clear();
        dblCopperLossesTransformer_MWh.Clear();
        intAllocationTransformer.Clear();
        dblEnergyTransformerBalanceBefore_MWh.Clear();
        dblProportionTransformer_pu.Clear();
        dblEnergyTransformerBalanceAfter_MWh.Clear();
        dblCorretionFatorTransformer_pu.Clear();
        dblCopperLossesCorrectedTransformer_MWh.Clear();
        dblEnergyTransformerResult_MWh.Clear();
        intFlow.Clear();
    }

    public void addTransformer(string strPrimSegmet, string strSecuSegmet, string strTercSegmet, double dblEnergy_MWh, double dblIronLosses_MWh, double dblCopperLosses_MWh, string strAllocation, int intFlow)
    {
        if (intPrimTransformer.Count != 0)
        {
            for (int i = 0; i < intPrimTransformer.Count; i++)
            {
                if (intFlow == 0)
                {
                    if (getLevelNumber(strPrimSegmet) == (int)intPrimTransformer[i] && getLevelNumber(strSecuSegmet) == (int)intSecuTransformer[i] && getLevelNumber(strTercSegmet) == (int)intTercTransformer[i] && getLevelNumber(strAllocation) == (int)intAllocationTransformer[i] && (dblEnergy_MWh != 0 || dblIronLosses_MWh != 0 || dblCopperLosses_MWh != 0))
                    {
                        dblEnergyTransformer_MWh[i] = (double)dblEnergyTransformer_MWh[i] + dblEnergy_MWh;
                        dblIronLossesTransformer_MWh[i] = (double)dblIronLossesTransformer_MWh[i] + dblIronLosses_MWh;
                        dblCopperLossesTransformer_MWh[i] = (double)dblCopperLossesTransformer_MWh[i] + dblCopperLosses_MWh;
                        return;
                    }
                }
                else
                {
                    if (getLevelNumber(strSecuSegmet) == (int)intPrimTransformer[i] && getLevelNumber(strPrimSegmet) == (int)intSecuTransformer[i] && getLevelNumber(strTercSegmet) == (int)intTercTransformer[i] && getLevelNumber(strAllocation) == (int)intAllocationTransformer[i] && (dblEnergy_MWh != 0 || dblIronLosses_MWh != 0 || dblCopperLosses_MWh != 0))
                    {
                        dblEnergyTransformer_MWh[i] = (double)dblEnergyTransformer_MWh[i] + dblEnergy_MWh;
                        dblIronLossesTransformer_MWh[i] = (double)dblIronLossesTransformer_MWh[i] + dblIronLosses_MWh;
                        dblCopperLossesTransformer_MWh[i] = (double)dblCopperLossesTransformer_MWh[i] + dblCopperLosses_MWh;
                        return;
                    }
                }
            }
        }
        if (dblEnergy_MWh != 0 || dblIronLosses_MWh != 0 || dblCopperLosses_MWh != 0)
        {
            if (intFlow == 0)
            {
                intPrimTransformer.Add(getLevelNumber(strPrimSegmet));
                intSecuTransformer.Add(getLevelNumber(strSecuSegmet));
            }
            else
            {
                intPrimTransformer.Add(getLevelNumber(strSecuSegmet));
                intSecuTransformer.Add(getLevelNumber(strPrimSegmet));
            }
            intTercTransformer.Add(getLevelNumber(strTercSegmet));
            dblEnergyTransformer_MWh.Add(dblEnergy_MWh);
            dblIronLossesTransformer_MWh.Add(dblIronLosses_MWh);
            dblCopperLossesTransformer_MWh.Add(dblCopperLosses_MWh);
            intAllocationTransformer.Add(getLevelNumber(strAllocation));
            dblEnergyTransformerBalanceBefore_MWh.Add(0.0);
            dblProportionTransformer_pu.Add(0.0);
            dblEnergyTransformerBalanceAfter_MWh.Add(0.0);
            dblCorretionFatorTransformer_pu.Add(0.0);
            dblCopperLossesCorrectedTransformer_MWh.Add(0.0);
            dblEnergyTransformerResult_MWh.Add(0.0);
            this.intFlow.Add(intFlow);
        }
    }

    private void updateProportionTransformer()
    {
        double dblEnergy_MWh = 0;
        for (int i = 0; i < intPrimTransformer.Count; i++)
        {
            dblEnergy_MWh = 0;
            for (int j = 0; j < intPrimTransformer.Count; j++)
            {
                if ((int)intPrimTransformer[i] == (int)intPrimTransformer[j] && (int)intSecuTransformer[i] == (int)intSecuTransformer[j] && (int)intTercTransformer[i] == (int)intTercTransformer[j])
                {
                    dblEnergy_MWh += (double)dblEnergyTransformer_MWh[j];
                }
            }
            dblProportionTransformer_pu[i] = (double)dblEnergyTransformer_MWh[i] / dblEnergy_MWh;
        }
    }

    private void updateTransformerBalance()
    {
        for (int i = 0; i < intPrimTransformer.Count; i++)
        {
            double dblEnergy_MWh = 0;
            dblEnergyTransformerBalanceBefore_MWh[i] = dblEnergy_MWh;
            for (int j = 0; j < intPrimSegmentTransformation.Count; j++)
            {
                for (int k = 0; k < intPrimSegmentTransformationSTtoT.Count; k++)
                {
                    if ((int)intPrimTransformer[i] == (int)intPrimTransformerSTtoT[k] && (int)intSecuTransformer[i] == (int)intSecuTransformerSTtoT[k] && (int)intTercTransformer[i] == (int)intTercTransformerSTtoT[k] && (int)intPrimSegmentTransformation[j] == (int)intPrimSegmentTransformationSTtoT[k] && (int)intSecuSegmentTransformation[j] == (int)intSecuSegmentTransformationSTtoT[k])
                    {
                        dblEnergy_MWh += (double)dblEnergySegmentTransformationBalanceAfter_MWh[j] * (double)dblProportionSTtoT_pu[k];
                    }
                }
            }
            dblEnergyTransformerBalanceBefore_MWh[i] = dblEnergy_MWh;
            dblEnergyTransformerBalanceAfter_MWh[i] = dblEnergy_MWh * (double)dblProportionTransformer_pu[i];
            //this.dblCorretionFatorTransformer_pu[i] = dblEnergy_MWh / (double)this.dblEnergyTransformer_MWh[i];
            dblCorretionFatorTransformer_pu[i] = (double)dblEnergyTransformerBalanceAfter_MWh[i] / (double)dblEnergyTransformer_MWh[i];
            dblCopperLossesCorrectedTransformer_MWh[i] = (double)dblCopperLossesTransformer_MWh[i] * Math.Pow((double)dblCorretionFatorTransformer_pu[i], 2);
            dblEnergyTransformerResult_MWh[i] = (double)dblEnergyTransformerBalanceBefore_MWh[i] - (double)dblIronLossesTransformer_MWh[i] - (double)dblCopperLossesCorrectedTransformer_MWh[i];
        }
    }

    // Energias transformadas
    private ArrayList intPrimSegmentTransformation = new ArrayList();

    private ArrayList intSecuSegmentTransformation = new ArrayList();
    private ArrayList dblEnergySegmentTransformation_MWh = new ArrayList();
    private ArrayList dblEnergySegmentTransformationBalanceBefore_MWh = new ArrayList();
    private ArrayList dblProportionSegmentTransformation_pu = new ArrayList();
    private ArrayList dblEnergySegmentTransformationBalanceAfter_MWh = new ArrayList();
    private ArrayList dblIronLossesSegmentTransformation_MWh = new ArrayList();
    private ArrayList dblCopperLossesSegmentTransformation_MWh = new ArrayList();
    private ArrayList dblEnergySegmentTransformationResult_MWh = new ArrayList();

    public void setSegmentTransformation()
    {
        intPrimSegmentTransformation.Clear();
        intSecuSegmentTransformation.Clear();
        dblEnergySegmentTransformation_MWh.Clear();
        dblEnergySegmentTransformationBalanceBefore_MWh.Clear();
        dblProportionSegmentTransformation_pu.Clear();
        dblEnergySegmentTransformationBalanceAfter_MWh.Clear();
        dblIronLossesSegmentTransformation_MWh.Clear();
        dblCopperLossesSegmentTransformation_MWh.Clear();
        dblEnergySegmentTransformationResult_MWh.Clear();
    }

    public void addSegmentTransformation(string strPrimSegmet, string strSecuSegmet, double dblEnergy_MWh)
    {
        if (intPrimSegmentTransformation.Count != 0)
        {
            for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
            {
                if (getLevelNumber(strPrimSegmet) == (int)intPrimSegmentTransformation[i] && getLevelNumber(strSecuSegmet) == (int)intSecuSegmentTransformation[i] && dblEnergy_MWh != 0)
                {
                    dblEnergySegmentTransformation_MWh[i] = (double)dblEnergySegmentTransformation_MWh[i] + dblEnergy_MWh;
                    return;
                }
            }
        }
        if (dblEnergy_MWh != 0)
        {
            intPrimSegmentTransformation.Add(getLevelNumber(strPrimSegmet));
            intSecuSegmentTransformation.Add(getLevelNumber(strSecuSegmet));
            dblEnergySegmentTransformation_MWh.Add(dblEnergy_MWh);
            dblEnergySegmentTransformationBalanceBefore_MWh.Add(0.0);
            dblProportionSegmentTransformation_pu.Add(0.0);
            dblEnergySegmentTransformationBalanceAfter_MWh.Add(0.0);
            dblIronLossesSegmentTransformation_MWh.Add(0.0);
            dblCopperLossesSegmentTransformation_MWh.Add(0.0);
            dblEnergySegmentTransformationResult_MWh.Add(0.0);
        }
    }

    private void updateProportionSegmentTransformation()
    {
        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            for (int j = 0; j < intPrimSegmentTransformation.Count; j++)
            {
                if ((int)intPrimSegmentTransformation[i] == (int)intPrimSegmentTransformation[j] && (int)intPrimSegmentTransformation[j] < (int)intSecuSegmentTransformation[j])
                {
                    dblProportionSegmentTransformation_pu[i] = (double)dblProportionSegmentTransformation_pu[i] + (double)dblEnergySegmentTransformation_MWh[j];
                }
            }
        }
        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            dblProportionSegmentTransformation_pu[i] = (double)dblEnergySegmentTransformation_MWh[i] / (double)dblProportionSegmentTransformation_pu[i];
            if ((int)intPrimSegmentTransformation[i] >= (int)intSecuSegmentTransformation[i])
            {
                dblProportionSegmentTransformation_pu[i] = (double)1;
            }
        }
    }

    private void updateEnergySegmentTransformationBalance()
    {
        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            if ((int)intPrimSegmentTransformation[i] < (int)intSecuSegmentTransformation[i])
            {
                for (int j = 0; j < inIsActiveLevel.GetLength(0); j++)
                {
                    if (inIsActiveLevel[j] == 1 && (int)intPrimSegmentTransformation[i] == j)
                    {
                        dblEnergySegmentTransformationBalanceBefore_MWh[i] = dblEnergyTransformerOutLevelUpDown_MWh[j];
                    }
                }
            }
            else if ((int)intPrimSegmentTransformation[i] >= (int)intSecuSegmentTransformation[i])
            {
                dblEnergySegmentTransformationBalanceBefore_MWh[i] = (double)dblEnergySegmentTransformation_MWh[i];
            }
        }

        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            dblEnergySegmentTransformationBalanceAfter_MWh[i] = (double)dblEnergySegmentTransformationBalanceBefore_MWh[i] * (double)dblProportionSegmentTransformation_pu[i];
        }
        /*
                                for (int i = 0; i < this.intPrimSegmentTransformation.Count; i++)
                                {
                                    double dblEnergy_MWh = 0;
                                    this.dblEnergySegmentTransformationBalanceAfter_MWh[i] = dblEnergy_MWh;
                                    for (int j = 0; j < this.intPrimSegmentTransformation.Count; j++)
                                    {
                                        if ((int)this.intPrimSegmentTransformation[i] == (int)this.intPrimSegmentTransformation[j] && (int)this.intSecuSegmentTransformation[i] == (int)this.intSecuSegmentTransformation[j])
                                        {
                                            dblEnergy_MWh += this.dblEnergyTransformerOutLevelUpDown_MWh[(int)this.intPrimSegmentTransformation[i]] * (double)this.dblProportionSegmentTransformation_pu[j];
                                        }
                                    }
                                    this.dblEnergySegmentTransformationBalanceAfter_MWh[i] = dblEnergy_MWh;
                                }
                    */
    }

    private void updateSegmentTransformationBalance()
    {
        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            double dblIronLossesEnergy_MWh = 0;
            double dblCopperLossesEnergy_MWh = 0;
            dblIronLossesSegmentTransformation_MWh[i] = dblIronLossesEnergy_MWh;
            dblCopperLossesSegmentTransformation_MWh[i] = dblCopperLossesEnergy_MWh;
            for (int j = 0; j < intPrimTransformer.Count; j++)
            {
                for (int k = 0; k < intPrimSegmentTransformationTtoST.Count; k++)
                {
                    if ((int)intPrimTransformer[j] == (int)intPrimTransformerTtoST[k] && (int)intSecuTransformer[j] == (int)intSecuTransformerTtoST[k] && (int)intTercTransformer[j] == (int)intTercTransformerTtoST[k] && (int)intPrimSegmentTransformation[i] == (int)intPrimSegmentTransformationTtoST[k] && (int)intSecuSegmentTransformation[i] == (int)intSecuSegmentTransformationTtoST[k])
                    {
                        dblIronLossesEnergy_MWh += (double)dblIronLossesTransformer_MWh[j] * (double)dblProportionTtoST_pu[k];
                        dblCopperLossesEnergy_MWh += (double)dblCopperLossesCorrectedTransformer_MWh[j] * (double)dblProportionTtoST_pu[k];
                    }
                }
            }
            dblIronLossesSegmentTransformation_MWh[i] = dblIronLossesEnergy_MWh;
            dblCopperLossesSegmentTransformation_MWh[i] = dblCopperLossesEnergy_MWh;
            dblEnergySegmentTransformationResult_MWh[i] = (double)dblEnergySegmentTransformationBalanceAfter_MWh[i] - dblIronLossesEnergy_MWh - dblCopperLossesEnergy_MWh;
        }
    }

    // Proporções que transformam dados de transformadores em energia transformadas
    private ArrayList intPrimTransformerTtoST = new ArrayList();

    private ArrayList intSecuTransformerTtoST = new ArrayList();
    private ArrayList intTercTransformerTtoST = new ArrayList();
    private ArrayList intPrimSegmentTransformationTtoST = new ArrayList();
    private ArrayList intSecuSegmentTransformationTtoST = new ArrayList();
    private ArrayList dblEnergyTtoST_MWh = new ArrayList();
    private ArrayList dblEnergyTotalTtoST_MWh = new ArrayList();
    private ArrayList dblProportionTtoST_pu = new ArrayList();

    public void setProportionTtoST()
    {
        intPrimTransformerTtoST.Clear();
        intSecuTransformerTtoST.Clear();
        intTercTransformerTtoST.Clear();
        intPrimSegmentTransformationTtoST.Clear();
        intSecuSegmentTransformationTtoST.Clear();
        dblEnergyTtoST_MWh.Clear();
        dblEnergyTotalTtoST_MWh.Clear();
        dblProportionTtoST_pu.Clear();
    }

    public void addProportionTtoST(string strPrimSegmetT, string strSecuSegmetT, string strTercSegmetT, string strPrimSegmetET, string strSecuSegmetET, double dblEnergyET_MWh)
    {
        if (intPrimTransformerTtoST.Count != 0)
        {
            for (int i = 0; i < intPrimTransformerTtoST.Count; i++)
            {
                if (getLevelNumber(strPrimSegmetT) == (int)intPrimTransformerTtoST[i] && getLevelNumber(strSecuSegmetT) == (int)intSecuTransformerTtoST[i] && getLevelNumber(strTercSegmetT) == (int)intTercTransformerTtoST[i] && getLevelNumber(strPrimSegmetET) == (int)intPrimSegmentTransformationTtoST[i] && getLevelNumber(strSecuSegmetET) == (int)intSecuSegmentTransformationTtoST[i] && dblEnergyET_MWh != 0)
                {
                    dblEnergyTtoST_MWh[i] = (double)dblEnergyTtoST_MWh[i] + dblEnergyET_MWh;
                    dblEnergyTotalTtoST_MWh[i] = 0.0;
                    dblProportionTtoST_pu[i] = 0.0;
                    return;
                }
            }
        }
        if (dblEnergyET_MWh != 0)
        {
            intPrimTransformerTtoST.Add(getLevelNumber(strPrimSegmetT));
            intSecuTransformerTtoST.Add(getLevelNumber(strSecuSegmetT));
            intTercTransformerTtoST.Add(getLevelNumber(strTercSegmetT));
            intPrimSegmentTransformationTtoST.Add(getLevelNumber(strPrimSegmetET));
            intSecuSegmentTransformationTtoST.Add(getLevelNumber(strSecuSegmetET));
            dblEnergyTtoST_MWh.Add(dblEnergyET_MWh);
            dblEnergyTotalTtoST_MWh.Add(0.0);
            dblProportionTtoST_pu.Add(0.0);
        }
    }

    private void updateEnergyTotalTtoST()
    {
        for (int i = 0; i < intPrimTransformerTtoST.Count; i++)
        {
            for (int j = 0; j < intPrimTransformerTtoST.Count; j++)
            {
                if ((int)intPrimTransformerTtoST[i] == (int)intPrimTransformerTtoST[j] && (int)intSecuTransformerTtoST[i] == (int)intSecuTransformerTtoST[j] && (int)intTercTransformerTtoST[i] == (int)intTercTransformerTtoST[j])
                {
                    dblEnergyTotalTtoST_MWh[i] = (double)dblEnergyTotalTtoST_MWh[i] + (double)dblEnergyTtoST_MWh[j];
                }
            }
        }
        updateProportionTtoST();
    }

    private void updateProportionTtoST()
    {
        for (int i = 0; i < intPrimTransformerTtoST.Count; i++)
        {
            dblProportionTtoST_pu[i] = (double)dblEnergyTtoST_MWh[i] / (double)dblEnergyTotalTtoST_MWh[i];
        }
    }

    // Proporções que transformam dados de energia transformadas em transformadores
    private ArrayList intPrimSegmentTransformationSTtoT = new ArrayList();

    private ArrayList intSecuSegmentTransformationSTtoT = new ArrayList();
    private ArrayList intPrimTransformerSTtoT = new ArrayList();
    private ArrayList intSecuTransformerSTtoT = new ArrayList();
    private ArrayList intTercTransformerSTtoT = new ArrayList();
    private ArrayList dblEnergySTtoT_MWh = new ArrayList();
    private ArrayList dblEnergyTotalSTtoT_MWh = new ArrayList();
    private ArrayList dblProportionSTtoT_pu = new ArrayList();

    public void setProportionSTtoT()
    {
        intPrimTransformerSTtoT.Clear();
        intSecuTransformerSTtoT.Clear();
        intTercTransformerSTtoT.Clear();
        intPrimSegmentTransformationSTtoT.Clear();
        intSecuSegmentTransformationSTtoT.Clear();
        dblEnergySTtoT_MWh.Clear();
        dblEnergyTotalSTtoT_MWh.Clear();
        dblProportionSTtoT_pu.Clear();
    }

    public void addProportionSTtoT(string strPrimSegmetET, string strSecuSegmetET, string strPrimSegmetT, string strSecuSegmetT, string strTercSegmetT, double dblEnergyT_MWh)
    {
        if (intPrimSegmentTransformationSTtoT.Count != 0)
        {
            for (int i = 0; i < intPrimSegmentTransformationSTtoT.Count; i++)
            {
                if (getLevelNumber(strPrimSegmetET) == (int)intPrimSegmentTransformationSTtoT[i] && getLevelNumber(strSecuSegmetET) == (int)intSecuSegmentTransformationSTtoT[i] && getLevelNumber(strPrimSegmetT) == (int)intPrimTransformerSTtoT[i] && getLevelNumber(strSecuSegmetT) == (int)intSecuTransformerSTtoT[i] && getLevelNumber(strTercSegmetT) == (int)intTercTransformerSTtoT[i] && dblEnergyT_MWh != 0)
                {
                    dblEnergySTtoT_MWh[i] = (double)dblEnergySTtoT_MWh[i] + dblEnergyT_MWh;
                    dblEnergyTotalSTtoT_MWh[i] = 0.0;
                    return;
                }
            }
        }
        if (dblEnergyT_MWh != 0)
        {
            intPrimSegmentTransformationSTtoT.Add(getLevelNumber(strPrimSegmetET));
            intSecuSegmentTransformationSTtoT.Add(getLevelNumber(strSecuSegmetET));
            intPrimTransformerSTtoT.Add(getLevelNumber(strPrimSegmetT));
            intSecuTransformerSTtoT.Add(getLevelNumber(strSecuSegmetT));
            intTercTransformerSTtoT.Add(getLevelNumber(strTercSegmetT));
            dblEnergySTtoT_MWh.Add(dblEnergyT_MWh);
            dblEnergyTotalSTtoT_MWh.Add(0.0);
            dblProportionSTtoT_pu.Add(0.0);
        }
    }

    private void updateEnergyTotalSTtoT()
    {
        for (int i = 0; i < intPrimSegmentTransformationSTtoT.Count; i++)
        {
            for (int j = 0; j < intPrimSegmentTransformationSTtoT.Count; j++)
            {
                if ((int)intPrimSegmentTransformationSTtoT[i] == (int)intPrimSegmentTransformationSTtoT[j] && (int)intSecuSegmentTransformationSTtoT[i] == (int)intSecuSegmentTransformationSTtoT[j])
                {
                    dblEnergyTotalSTtoT_MWh[i] = (double)dblEnergyTotalSTtoT_MWh[i] + (double)dblEnergySTtoT_MWh[j];
                }
            }
        }
        updateProportionSTtoT();
    }

    private void updateProportionSTtoT()
    {
        for (int i = 0; i < intPrimSegmentTransformationSTtoT.Count; i++)
        {
            dblProportionSTtoT_pu[i] = (double)dblEnergySTtoT_MWh[i] / (double)dblEnergyTotalSTtoT_MWh[i];
        }
    }

    // Energia de transformadores saindo do nível destinado a nível superior
    private double[] dblEnergyTransformerOutLevelDownUp_MWh = new double[5];

    private void setEnergyTransformerOutLevelDownUp_MWh()
    {
        for (int i = 0; i < dblEnergyTransformerOutLevelDownUp_MWh.GetLength(0); i++)
        {
            dblEnergyTransformerOutLevelDownUp_MWh[i] = 0;
        }
    }

    private void setEnergyTransformerOutLevelDownUp_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblEnergyTransformerOutLevelDownUp_MWh[intSegmet] = 0.0;
            for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
            {
                if (intSegmet == (int)intPrimSegmentTransformation[i] && (int)intPrimSegmentTransformation[i] > (int)intSecuSegmentTransformation[i])
                {
                    dblEnergyTransformerOutLevelDownUp_MWh[intSegmet] = dblEnergyTransformerOutLevelDownUp_MWh[intSegmet] + (double)dblEnergySegmentTransformationBalanceAfter_MWh[i];
                }
            }
        }
    }

    // Energia de transformadores entrando no nível proveniente de nível inferior
    private double[] dblEnergyTransformerEnterLevelDownUp_MWh = new double[5];

    private void setEnergyTransformerEnterLevelDownUp_MWh()
    {
        for (int i = 0; i < dblEnergyTransformerEnterLevelDownUp_MWh.GetLength(0); i++)
        {
            dblEnergyTransformerEnterLevelDownUp_MWh[i] = 0;
        }
    }

    private void setEnergyTransformerEnterLevelDownUp_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblEnergyTransformerEnterLevelDownUp_MWh[intSegmet] = 0.0;
            for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
            {
                if (intSegmet == (int)intSecuSegmentTransformation[i] && (int)intPrimSegmentTransformation[i] > (int)intSecuSegmentTransformation[i])
                {
                    dblEnergyTransformerEnterLevelDownUp_MWh[intSegmet] = dblEnergyTransformerEnterLevelDownUp_MWh[intSegmet] + (double)dblEnergySegmentTransformationResult_MWh[i];
                }
            }
        }
    }

    // Energia de transformadores entrando no nível proveniente de nível superior
    private double[] dblEnergyTransformerEnterLevelUpDown_MWh = new double[5];

    private void setEnergyTransformerEnterLevelUpDown_MWh()
    {
        for (int i = 0; i < dblEnergyTransformerEnterLevelUpDown_MWh.GetLength(0); i++)
        {
            dblEnergyTransformerEnterLevelUpDown_MWh[i] = 0;
        }
    }

    private void setEnergyTransformerEnterLevelUpDown_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblEnergyTransformerEnterLevelUpDown_MWh[intSegmet] = 0.0;
            for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
            {
                if (intSegmet == (int)intSecuSegmentTransformation[i] && (int)intPrimSegmentTransformation[i] < (int)intSecuSegmentTransformation[i])
                {
                    dblEnergyTransformerEnterLevelUpDown_MWh[intSegmet] = dblEnergyTransformerEnterLevelUpDown_MWh[intSegmet] + (double)dblEnergySegmentTransformationResult_MWh[i];
                }
            }
        }
    }

    // Energia de transformadores saindo do nível destinado a nível inferior
    private double[] dblEnergyTransformerOutLevelUpDown_MWh = new double[5];

    private void setEnergyTransformerOutLevelUpDown_MWh()
    {
        for (int i = 0; i < dblEnergyTransformerOutLevelUpDown_MWh.GetLength(0); i++)
        {
            dblEnergyTransformerOutLevelUpDown_MWh[i] = 0;
        }
    }

    private void setEnergyTransformerOutLevelUpDown_MWh(int intSegmet)
    {
        if (intSegmet != -1)
        {
            dblEnergyTransformerOutLevelUpDown_MWh[intSegmet] = dblBalanceLevelEnergy_MWh[intSegmet] - (dblTecnicalLossesFixedEnergy_MWh[intSegmet] - dblTecnicalLossesTransformersEnergy_MWh[intSegmet]) - dblSupplyEnergywithNet_MWh[intSegmet];
        }
    }

    private void setProportion()
    {
        /*
        for (int i = 0; i < this.strPrimTransformer.Count; i++)
        {
            for (int j = 0; j < this.strPrimSegmentTransformation.Count; j++)
            {
                if ((strPrimTransformer[i].ToString() == strPrimSegmentTransformation[j].ToString() && strSecuTransformer[i].ToString() == strSecuSegmentTransformation[j].ToString()) || (strPrimTransformer[i].ToString() == strPrimSegmentTransformation[j].ToString() && strTercTransformer[i].ToString() == strSecuSegmentTransformation[j].ToString()))
                {
                    this.strPrimTransformerTtoST.Add(strPrimTransformer[i]);
                    this.strSecuTransformerTtoST.Add(strSecuTransformer[i]);
                    this.strTercTransformerTtoST.Add(strTercTransformer[i]);
                    this.strPrimSegmentTransformationTtoST.Add(strPrimSegmentTransformation[j]);
                    this.strSecuSegmentTransformationTtoST.Add(strSecuSegmentTransformation[j]);
                    this.dblProportionTtoST_pu.Add((double[])this.dblEnergySegmentTransformation_MWh[j]);
                }
            }
        }

        ArrayList strPrimST = new ArrayList();
        ArrayList dblEnergy_MWh = new ArrayList();
        for (int i = 0; i < this.strPrimSegmentTransformationTtoST.Count; i++)
        {
            if (!strPrimST.Contains(this.strPrimSegmentTransformationTtoST[i]))
            {
                strPrimST.Add(this.strPrimSegmentTransformationTtoST[i]);
                dblEnergy_MWh.Add((double[])this.dblProportionTtoST_pu[i]);
            }
            else
            {
                int index = strPrimST.IndexOf(this.strPrimSegmentTransformationTtoST[i]);
                double[] values1 = (double[])dblEnergy_MWh[index];
                double[] values2 = (double[])this.dblProportionTtoST_pu[i];
                double[] values3 = {0,0,0,0,0,0,0,0,0,0,0,0};
                for (int j = 0; j <= 11; j++)
                    values3[j] = values1[j]+values2[j];
                dblEnergy_MWh[index] = values3;
            }
        }

        for (int i = 0; i < this.strPrimSegmentTransformationTtoST.Count; i++)
        {
            int index = strPrimST.IndexOf(this.strPrimSegmentTransformationTtoST[i]);
            double[] values1 = (double[])this.dblProportionTtoST_pu[i];
            double[] values2 = (double[])dblEnergy_MWh[index];
            double[] values3 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            for (int j = 0; j <= 11; j++)
                values3[j] = values1[j] / values2[j];
            this.dblProportionTtoST_pu[i] = values3;
        }
        */
    }

    public void runBalance()
    {
        // Verifica que níveis estão ativos
        setIsActiveLevel();

        // Calcula as proporções da tabela de energias transformadas
        updateProportionSegmentTransformation();

        // Calcula as energias totais e proporções dos Transformadores para Energias Transformadas
        updateEnergyTotalTtoST();

        // Calcula as energias totais e proporções das Energias Transformadas para Transformadores
        updateEnergyTotalSTtoT();

        // Calcula a energia fornecida ao consumidores que transita na rede
        setSupplyEnergywithNet_MWh();

        // Calcula a energia passante do balanço nos níveis
        setBalanceLevelEnergy_MWh();

        // Calcula a energia que deverá sair do nível para baixo
        setEnergyTransformerOutLevelUpDown_MWh();

        // Calcula a energia que entra nos segmentos de transformação proveniente do balanço
        updateEnergySegmentTransformationBalance();

        // Calcula a diferença dos níveis
        setDiferenceBalanceEnergy_MWh();

        // Não tem mais utilidade
        //this.setProportion();

        for (int i = 0; i <= 100; i++)
        {
            for (int j = 0; j <= 4; j++)
            {
                if (inIsActiveLevel[j] == 1)
                {
                    setEnergyTransformerOutLevelDownUp_MWh(j);
                    setEnergyTransformerEnterLevelDownUp_MWh(j);
                    setTecnicalLossesTransformersEnergy_MWh(j);
                    setBalanceLevelEnergy_MWh(j);
                    setEnergyTransformerOutLevelUpDown_MWh(j);
                    updateEnergySegmentTransformationBalance();
                    setDiferenceBalanceEnergy_MWh(j);
                    updateProportionTransformer();
                    updateTransformerBalance();
                    updateSegmentTransformationBalance();
                    setEnergyTransformerEnterLevelUpDown_MWh(j);
                }
            }
        }
    }

    public static int getLevelNumber(string strSegment)
    {
        int intLevel = -1;

        if (strSegment == "A1")
            intLevel = 0;
        else if (strSegment == "A2")
            intLevel = 1;
        else if (strSegment == "A3")
            intLevel = 2;
        else if (strSegment == "A3a")
            intLevel = 3;
        else if (strSegment == "A4")
            intLevel = 3;
        else if (strSegment == "MT")
            intLevel = 3;
        else if (strSegment == "B")
            intLevel = 4;

        return intLevel;
    }

    public static string getLevelSegment(int intSegment)
    {
        string strLevel = "NA";

        if (intSegment == 0)
            strLevel = "A1";
        else if (intSegment == 1)
            strLevel = "A2";
        else if (intSegment == 2)
            strLevel = "A3";
        else if (intSegment == 3)
            strLevel = "MT";
        else if (intSegment == 4)
            strLevel = "B";

        return strLevel;
    }

    public string ToString()
    {
        string strResult = "";

        strResult += "Segmento\tEnergia Injetada (MWh)\tEnergia sem Rede Associada (MWh)\tEnergia Proveniente de Níveis Inferiores (MWh)\tEnergia Destinada a Níveis Superiores (MWh)\tEnergia Proveniente de Níveis Superiores (MWh)\tEnergia Base do Nível (MWh)\tPerda Técnica de Energia Não Corrigível (MWh)\tPerda Técnica de Energia Relativa aos Transformadores (MWh)\tEnergia Destinada a Níveis Inferiores (MWh)\tEnergia Fornecida Descontada Energia sem Rede Associada (MWh)\tDifença (MWh)" + Environment.NewLine;
        for (int i = 0; i <= 4; i++)
        {
            if (inIsActiveLevel[i] == 1)
            {
                strResult += getLevelSegment(i);
                strResult += "\t";
                strResult += dblInjectEnergy_MWh[i].ToString();
                strResult += "\t";
                strResult += dblSupplyEnergywithoutNet_MWh[i].ToString();
                strResult += "\t";
                strResult += dblEnergyTransformerEnterLevelDownUp_MWh[i].ToString();
                strResult += "\t";
                strResult += dblEnergyTransformerOutLevelDownUp_MWh[i].ToString();
                strResult += "\t";
                strResult += dblEnergyTransformerEnterLevelUpDown_MWh[i].ToString();
                strResult += "\t";
                strResult += dblBalanceLevelEnergy_MWh[i].ToString();
                strResult += "\t";
                strResult += dblTecnicalLossesFixedEnergy_MWh[i].ToString();
                strResult += "\t";
                strResult += dblTecnicalLossesTransformersEnergy_MWh[i].ToString();
                strResult += "\t";
                strResult += dblEnergyTransformerOutLevelUpDown_MWh[i].ToString();
                strResult += "\t";
                strResult += dblSupplyEnergywithNet_MWh[i].ToString();
                strResult += "\t";
                strResult += dblDiferenceBalanceEnergy_MWh[i].ToString();
                strResult += Environment.NewLine;
            }
        }

        strResult += "Segmento\tEnergia dos Transformadores (MWh)\tPerda de Energia Ferro (MWh)\tPerda de Energia Cobre (MWh)\tNível de Alocação\tEnergia do Balanço Originada da ET (MWh)\tProporção (p.u.)\tEnergia do Balanço Originada da ET Corrigida (MWh)\tFator de Correção (p.u.)\tPerda de Energia Cobre Corrigido (MWh)\tEnergia do Resultante (MWh)" + Environment.NewLine;
        for (int i = 0; i < intPrimTransformer.Count; i++)
        {
            strResult += getLevelSegment((int)intPrimTransformer[i]) + "-" + getLevelSegment((int)intSecuTransformer[i]) + "-" + getLevelSegment((int)intTercTransformer[i]);
            strResult += "\t";
            strResult += ((double)dblEnergyTransformer_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblIronLossesTransformer_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblCopperLossesTransformer_MWh[i]).ToString();
            strResult += "\t";
            strResult += getLevelSegment((int)intAllocationTransformer[i]);
            strResult += "\t";
            strResult += ((double)dblEnergyTransformerBalanceBefore_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblProportionTransformer_pu[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblEnergyTransformerBalanceAfter_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblCorretionFatorTransformer_pu[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblCopperLossesCorrectedTransformer_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblEnergyTransformerResult_MWh[i]).ToString();
            strResult += Environment.NewLine;
        }

        strResult += "Segmento\tEnergia dos Segmentos de Transformação (MWh)\tEnergia da Barra dos Níveis (MWh)\tProporções (p.u.)\tEnergia Injetada Transformador do Balanço (MWh)\tPerda de Energia Ferro (MWh)\tPerda de Energia Cobre (MWh)\tEnergia do Resultante Transformador (MWh)" + Environment.NewLine;
        for (int i = 0; i < intPrimSegmentTransformation.Count; i++)
        {
            strResult += getLevelSegment((int)intPrimSegmentTransformation[i]) + "-" + getLevelSegment((int)intSecuSegmentTransformation[i]);
            strResult += "\t";
            strResult += ((double)dblEnergySegmentTransformation_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblEnergySegmentTransformationBalanceBefore_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblProportionSegmentTransformation_pu[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblEnergySegmentTransformationBalanceAfter_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblIronLossesSegmentTransformation_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblCopperLossesSegmentTransformation_MWh[i]).ToString();
            strResult += "\t";
            strResult += ((double)dblEnergySegmentTransformationResult_MWh[i]).ToString();
            strResult += Environment.NewLine;
        }

        strResult += "Transformador\tSegmento de Transformação\tProporção (p.u.)" + Environment.NewLine;
        for (int i = 0; i < intPrimTransformerTtoST.Count; i++)
        {
            strResult += getLevelSegment((int)intPrimTransformerTtoST[i]) + "-" + getLevelSegment((int)intSecuTransformerTtoST[i]) + "-" + getLevelSegment((int)intTercTransformerTtoST[i]);
            strResult += "\t";
            strResult += getLevelSegment((int)intPrimSegmentTransformationTtoST[i]) + "-" + getLevelSegment((int)intSecuSegmentTransformationTtoST[i]);
            strResult += "\t";
            strResult += ((double)dblProportionTtoST_pu[i]).ToString();
            strResult += Environment.NewLine;
        }

        strResult += "Segmento de Transformação\tTransformador\tProporção (p.u.)" + Environment.NewLine;
        for (int i = 0; i < intPrimTransformerSTtoT.Count; i++)
        {
            strResult += getLevelSegment((int)intPrimSegmentTransformationSTtoT[i]) + "-" + getLevelSegment((int)intSecuSegmentTransformationSTtoT[i]);
            strResult += "\t";
            strResult += getLevelSegment((int)intPrimTransformerSTtoT[i]) + "-" + getLevelSegment((int)intSecuTransformerSTtoT[i]) + "-" + getLevelSegment((int)intTercTransformerSTtoT[i]);
            strResult += "\t";
            strResult += ((double)dblProportionSTtoT_pu[i]).ToString();
            strResult += Environment.NewLine;
        }

        return strResult;
    }
}