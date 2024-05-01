namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.OpenDss;

public class DssWrapper
{    
    
    public Circuit Circuit { get; set; }
    public Solution Solution { get; set; }
    public CtrlQueue SolutionControlQueue { get; set; }
    public DSS DSSServer { get; set; }
    public Text? Text { get; set; }

    public Type Meters
    {
        get
        {
            return Meters.GetType();
        }
    }

    public string getSequenceVoltages()
    {
        Bus dssBus;
        int i, j;
        dynamic V, cpxV;
        string strTextoFull = "";

        for (i = 0; i < Circuit.NumBuses; i++)
        {
            dssBus = Circuit.Buses[i];
            var strTexto = "\t" + Circuit.ActiveBus.Name;

            V = dssBus.SeqVoltages;
            cpxV = dssBus.CplxSeqVoltages;

            for (j = 0; j < V.Length; j++)
            {
                strTexto += "\t" + V[j];
            }

            for (j = 0; j < cpxV.Length; j++)
            {
                strTexto += "\t" + cpxV[j];
            }
            strTextoFull += strTexto;
        }
        return strTextoFull;
    }

    public string getVoltages()
    {
        Bus dssBus;
        int i, j;
        dynamic V;
        string strTextoFull = "";

        for (i = 0; i < Circuit.NumBuses; i++)
        {
            dssBus = Circuit.Buses[i];
            var strTexto = "\t" + Circuit.ActiveBus.Name;

            V = dssBus.puVoltages;

            for (j = 0; j < V.Length; j++)
            {
                strTexto += "\t" + V[j];
            }

            strTextoFull += strTexto;
        }
        return strTextoFull;
    }

    public string getMagnitudeAngle()
    {
        Bus dssBus;
        int i, j;
        dynamic V;
        string strTextoFull = "";

        for (i = 0; i < Circuit.NumBuses; i++)
        {
            dssBus = Circuit.Buses[i];
            var strTexto = "\t" + Circuit.ActiveBus.Name;

            V = dssBus.puVoltages;

            for (j = 0; j < V.Length; j = j + 2)
            {
                strTexto += "\t" + "(" + Math.Sqrt(V[j] * V[j] + V[j + 1] * V[j + 1]) + "; " + angleComplex(V[j], V[j + 1]) + ")";
            }

            strTextoFull += strTexto;
        }
        return strTextoFull;
    }

    public string getSequenceCurrents()
    {
        CktElement dssElement;
        int i, j;
        dynamic SeqCurr, cpxSeqValues;
        string strTextoFull = "";

        dssElement = Circuit.ActiveCktElement;
        i = Circuit.FirstPDElement();

        while (i > 0)
        {
            if (dssElement.NumPhases == 3)
            {
                var strTexto = "\t" + dssElement.Name;

                SeqCurr = dssElement.SeqCurrents;
                cpxSeqValues = dssElement.CplxSeqCurrents;

                for (j = 0; j < SeqCurr.Length; j++)
                {
                    strTexto += "\t" + SeqCurr[j];
                }

                for (j = 0; j < cpxSeqValues.Length; j++)
                {
                    strTexto += "\t" + cpxSeqValues[j];
                }
                strTextoFull += strTexto;
            }
            i = Circuit.NextPDElement();
        }
        return strTextoFull;
    }

    public string getLoadPowers()
    {
        CktElement dssElement;
        int i, j;
        dynamic PCPowers;
        double SumkW, Sumkvar;
        string strTextoFull = "";

        dssElement = Circuit.ActiveCktElement;

        i = Circuit.Loads.First;

        while (i > 0)
        {
            var strTexto = "\t" + dssElement.Name;

            PCPowers = dssElement.Powers;

            SumkW = 0;
            Sumkvar = 0;

            for (j = 0; j < PCPowers.Length; j = j + 2)
            {
                SumkW = SumkW + PCPowers[j];
                Sumkvar = Sumkvar + PCPowers[j + 1];
            }

            strTexto += "\t" + "(" + SumkW + "; " + Sumkvar + ")";

            strTextoFull += strTexto;
            i = Circuit.Loads.Next;
        }
        return strTextoFull;
    }

    public static double angleComplex(double Re, double Im)
    {
        double angle;
        if (Re != 0)
        {
            angle = Math.Atan(Im / Re) * 180 / Math.PI;
            if (Re < 0)
            {
                angle = angle + 180;
                return angle;
            }
            else
                return angle;
        }
        else
        {
            if (Im > 0)
                return 90;
            else if (Im < 0)
                return -90;
            else
                return 0;
        }
    }

    public void InitiateServer()
    {
        this.DSSServer = new DSS();
    }
}