using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;

public enum MetodoParametros
{
    REGULATORIO,
    DISTRIBUIDORA
}

public static class FormConfigFactory
{
    public static FormConfigControls Create(MetodoParametros metodo)
    {
        if (metodo == MetodoParametros.REGULATORIO)
        {
            return ConfigRegulatorio();
        }
        else
        {
            return new FormConfigControls();
        }
    }

    private static FormConfigControls ConfigRegulatorio()
    {
        return new FormConfigControls
        {            
            clbOption = new Dictionary<string, bool>()
            {
                { "Convergência de Perda Não Técnica", true },
                { "Transformadores ABNT", true },
                { "Adequação Tensão Mínima das Cargas MT (0,93 pu)", true },
                { "Adequação Tensão Mínima das Cargas BT (0,92 pu)", true },
                { "Limitar Máxima Tensão de Barras e Reguladores (1,05 pu)", true },
                { "Limitar o Ramal (30m)", true },
                { "Utilizar Tap nos Transformadores", true },
                { "Limitar Cargas BT (Potência ativa do transformador)", true },
                { "Neutralizar Transformadores de Terceiros", true },
                { "Neutralizar Redes de Terceiros (MT/BT)", true },
                { "Eliminar Transformadores Vazios", true }
            },
            tbVPUMin = "0.6",
            cbModel = 1,
            cbReinicializeNT = true, 
            cbMeterComplete = true
        };
    }
}
