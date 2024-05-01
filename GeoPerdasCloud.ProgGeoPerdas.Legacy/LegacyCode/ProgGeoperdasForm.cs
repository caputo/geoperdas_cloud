using dss_sharp;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.OpenDss;
using System.IO.Compression;

namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;
public class ProgGeoperdasForm: IDisposable
{
    private string strCommandConnection;
    public SqlConnection sqlConnServer;
    private SqlCommand sqlCommServer;
    private SqlDataReader sqlDtRdrServer;

    private string strBaseSelected;
    private ArrayList strFeedersSelected = new ArrayList();

    private StreamWriter swFile;

    //private long lngindex = 0;
    private string[] strFileData = new string[1000000];


    private int intConverged = 0;

    private int intRealizaCnvrgcPNT;
    private int intUsaTrafoABNT;
    private int intAdequarTensaoCargasMT;
    private int intAdequarTensaoCargasBT;
    private int intAdequarTensaoSuperior;
    private int intAdequarRamal;
    private int intAdequarModeloCarga;
    private int intAdequarTapTrafo;
    private int intAdequarPotenciaCarga;
    private int intAdequarTrafoVazio;
    private int intNeutralizarTrafoTerceiros;
    private int intNeutralizarRedeTerceiros;

    private double dblVPUMin;

    private int intModeloConverge;

    private double dblEnergyInjectDeclared_MWh;
    private double dblEnergyLoadsMTDeclared_MWh;
    private double dblEnergyLoadsBTDeclared_MWh;
    private double dblEnergyLoadsMTNTCalculated_MWh;
    private double dblEnergyLoadsBTNTCalculated_MWh;
    private double dblPropEnergyLoadsMT_pu;
    private double dblPropEnergyLoadsBT_pu;

    private double dblEnergyInjectDeclaredYearly_MWh;
    private double dblEnergyInjectCalculatedYearly_MWh;

    private double dblTotalPower_kW;
    private double dblTotalEnergy_kWh;
    private double dblTecnicalLossPower_kW;
    private double dblTecnicalLossEnergy_kWh;
    private double dblTotalLoadEnergy_kWh;

    private double dblNoLoadLosses_kWh;
    private double dblLineLosses_kWh;

    private double dblTransformerLosses_kWh;
    private double dblA3aA3aTransformerLosses_kWh;
    private double dblA3aA4TransformerLosses_kWh;
    private double dblA4A3aTransformerLosses_kWh;
    private double dblA3aBTransformerLosses_kWh;
    private double dblBA3aTransformerLosses_kWh;
    private double dblA4A4TransformerLosses_kWh;
    private double dblA4BTransformerLosses_kWh;
    private double dblBA4TransformerLosses_kWh;
    private double dblA3aA3aTransformerNoLoadLosses_kWh;
    private double dblA3aA4TransformerNoLoadLosses_kWh;
    private double dblA4A3aTransformerNoLoadLosses_kWh;
    private double dblA3aBTransformerNoLoadLosses_kWh;
    private double dblBA3aTransformerNoLoadLosses_kWh;
    private double dblA4A4TransformerNoLoadLosses_kWh;
    private double dblA4BTransformerNoLoadLosses_kWh;
    private double dblBA4TransformerNoLoadLosses_kWh;

    private double dblMTLosses_kWh;
    private double dblA3aLosses_kWh;
    private double dblA4Losses_kWh;
    private double dblBTLosses_kWh;

    private double dblLineMTLosses_kWh;
    private double dblLineA3aLosses_kWh;
    private double dblLineA4Losses_kWh;
    private double dblLineBTLosses_kWh;

    private double dblMTLoad_kWh;
    private double dblBTLoad_kWh;

    private FormConfigControls config = FormConfigFactory.Create(MetodoParametros.REGULATORIO);
    private readonly ILogger<ProgGeoperdasForm> _logger;
    private DssWrapper _dss;

    public ProgGeoperdasForm(FormConfigControls pConfig)
    {
        config = pConfig; 
        using ILoggerFactory loggerFactory =
        LoggerFactory.Create(builder =>
        builder.AddSimpleConsole(options =>
        {
            options.IncludeScopes = true;
            options.SingleLine = true;
            options.TimestampFormat = "HH:mm:ss ";
        }));
        _logger = loggerFactory.CreateLogger<ProgGeoperdasForm>();
        _logger.LogInformation("Iniciado");
        mthdStarDSSServer();
    }
    public void Dispose()
    {
        mthdCloseDataBase();
        sqlDtRdrServer?.Dispose();
        sqlCommServer?.Dispose();
    }

    public void Close()
    {
        mthdCloseDataBase();
    }

    public void bConnection_Click()
    {
        if (config.tbServer == "" || config.tbServer == null)
        {
            mthdInputMensageLog("O nome do servidor não foi estabelecido.");
        }
        else if (config.tbDataBase == "" || config.tbDataBase == null)
        {
            mthdInputMensageLog("O nome da base de dados não foi estabelecida.");
        }
        else
        {
            mthdCloseDataBase();
            mthdOpenDataBase(config.tbServer, config.tbDataBase, config.tbUser, config.tbPassword);
        }
    }

    private void bDesconnection_Click()
    {
        mthdCloseDataBase();
    }

    public void bExecuteBD_Click()
    {
        int i = 0, intCtrlCnvrg = 0, intQuantIntrtn = 0;
        string strTempSqlCommand, strTensoesBase;
        int intAno = 0;
        double tolerance;
        double dblEnerInjDU_kWh = 0, dblEnerInjSA_kWh = 0, dblEnerInjDO_kWh = 0;
        double dblPerdaEnerTecDU_kWh = 0, dblPerdaEnerTecSA_kWh = 0, dblPerdaEnerTecDO_kWh = 0;
        double dblEnerForncDU_kWh = 0, dblEnerForncSA_kWh = 0, dblEnerForncDO_kWh = 0;
        double dblDelta = 0, dblDeltaMT = 0, dblDeltaBT = 0, dblDelta_MWh = 0, dblDeltaBT_MWh = 0, dblDeltaMT_MWh = 0, dblErrorEnerInj = 0;

        if (sqlConnServer == null)
        {
            mthdInputMensageLog("A conexão com o banco de dados não foi estabelecida.");
        }
        else if (config.tbObjectBase == null || config.tbObjectBase == null)
        {
            mthdInputMensageLog("A base a ser calculada não foi indicada.");
        }
        else
        {
            try
            {
                // Verifica as opções solicitadas pelo usuário
                mthdUpdateOptionsRun();

                // Captura a base a ser calculada
                strBaseSelected = config.tbObjectBase;

                // Captura o ano da base
                sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "Base] AS t1 WHERE t1.CodBase = " + strBaseSelected, sqlConnServer);
                sqlCommServer.CommandTimeout = 0;
                sqlDtRdrServer = sqlCommServer.ExecuteReader();
                sqlDtRdrServer.Read();
                intAno = Convert.ToInt32(sqlDtRdrServer["Ano"]);
                sqlDtRdrServer.Close();

                // Loop da base e dos alimentadores
                strFeedersSelected.Clear();
                if (config.tbObjectFeeder == null || config.tbObjectFeeder == "")
                    sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CircAtipResultante = 0", sqlConnServer);
                else
                    sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND " + GeoPerdasTools.mthdQuerySQLFeeders(config.tbObjectFeeder) + " AND t1.CircAtipResultante = 0", sqlConnServer);
                sqlCommServer.CommandTimeout = 0;
                sqlDtRdrServer = sqlCommServer.ExecuteReader();
                while (sqlDtRdrServer.Read())
                {
                    strFeedersSelected.Add(sqlDtRdrServer["CodAlim"].ToString());
                }
                sqlDtRdrServer.Close();

                foreach (string strFeeder in strFeedersSelected)
                {
                    try
                    {
                        // Modelo de convergência
                        // Modelo 1 corresponde ao modelo de convergência primário
                        // Modelo 2 corresponde ao modelo de convergência secundário
                        intModeloConverge = 1;

                        // Variáveis de avaliação de troca de modelo
                        dblEnergyInjectDeclaredYearly_MWh = 0;
                        dblEnergyInjectCalculatedYearly_MWh = 0;

                        // Início do processo
                        mthdInputMensageLog("Início do processo para o alimentador " + strFeeder + ".");

                        mthdStarCommands();

                        // Zera as cargas não técnicas calculadas anteriormente para reiniciar o processo
                        if (config.cbReinicializeNT == true)
                        {
                            sqlCommServer = new SqlCommand("dbo.iadAuxCargaBTNTDem", sqlConnServer);
                            sqlCommServer.CommandType = CommandType.StoredProcedure;
                            sqlCommServer.Parameters.AddWithValue("@CodBase", strBaseSelected);
                            sqlCommServer.Parameters.AddWithValue("@CodOperation", 1);
                            sqlCommServer.Parameters.AddWithValue("@CodAlim", strFeeder);
                            sqlCommServer.ExecuteNonQuery();

                            sqlCommServer = new SqlCommand("dbo.iadAuxCargaMTNTDem", sqlConnServer);
                            sqlCommServer.CommandType = CommandType.StoredProcedure;
                            sqlCommServer.Parameters.AddWithValue("@CodBase", strBaseSelected);
                            sqlCommServer.Parameters.AddWithValue("@CodOperation", 1);
                            sqlCommServer.Parameters.AddWithValue("@CodAlim", strFeeder);
                            sqlCommServer.ExecuteNonQuery();
                        }

                        // Cria estrutura de tabelas para convergência não técnica
                        if (intRealizaCnvrgcPNT == 1)
                        {
                            sqlCommServer = new SqlCommand("dbo.CriaTabelaDemandaCargaNTAlimentador", sqlConnServer);
                            sqlCommServer.CommandType = CommandType.StoredProcedure;
                            sqlCommServer.Parameters.AddWithValue("@BaseAnalisada", strBaseSelected);
                            sqlCommServer.Parameters.AddWithValue("@AlimAnalisado", strFeeder);
                            sqlCommServer.ExecuteNonQuery();
                        }

                        // Início do código
                        mthdInsertCommandInLine("Clear");

                        // Loop dos códigos dos circuitos
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de circuitos");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandCircuit(strFeeder, Convert.ToDouble(sqlDtRdrServer["TenNom_kV"]), Convert.ToDouble(sqlDtRdrServer["TenOpe_pu"]), sqlDtRdrServer["CodPonAcopl"].ToString(), intAdequarTensaoSuperior);
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelo circuito.");

                        // Loop dos códigos dos condutores
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "CodCondutor] AS t1 WHERE t1.CodBase = " + strBaseSelected, sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de condutores");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLinecode(sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(sqlDtRdrServer["Resis_ohms_km"]), Convert.ToDouble(sqlDtRdrServer["Reat_ohms_km"]), Convert.ToDouble(sqlDtRdrServer["CorrMax_A"]));
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos condutores.");

                        // Loop dos códigos das curvas MT
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxCrvCrgMTHor AS t1 WHERE t1.CodBase = " + strBaseSelected, sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de curvas MT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLoadshape(sqlDtRdrServer["CodCrvCrg"].ToString() + "_" + sqlDtRdrServer["TipoDia"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm01"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm02"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm03"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm04"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm05"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm06"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm07"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm08"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm09"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm10"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm11"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm12"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm13"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm14"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm15"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm16"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm17"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm18"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm19"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm20"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm21"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm22"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm23"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm24"]));
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelas curvas MT.");

                        // Loop dos códigos das curvas BT
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxCrvCrgBTHor AS t1 WHERE t1.CodBase = " + strBaseSelected, sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de curvas BT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLoadshape(sqlDtRdrServer["CodCrvCrg"].ToString() + "_" + sqlDtRdrServer["TipoDia"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm01"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm02"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm03"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm04"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm05"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm06"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm07"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm08"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm09"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm10"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm11"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm12"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm13"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm14"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm15"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm16"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm17"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm18"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm19"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm20"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm21"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm22"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm23"]), Convert.ToDouble(sqlDtRdrServer["PotAtvNorm24"]));
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelas curvas BT.");

                        // Loop das chaves MT
                        sqlCommServer = new SqlCommand("SELECT t1.CodChvMT, t1.CodFas, t1.De, t1.Para FROM dbo.[" + strBaseSelected + "ChaveMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "' AND t1.EstChv = 2", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de chaves MT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandMTSwitch("CMT_" + sqlDtRdrServer["CodChvMT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["De"].ToString(), sqlDtRdrServer["Para"].ToString());
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelas chaves MT.");

                        // Loop dos segmentos MT
                        sqlCommServer = new SqlCommand("SELECT t1.CodSegmMT, t1.CodFas, t1.CodCond, t1.Comp_km, t1.De, t1.Para, t1.Propr FROM dbo.[" + strBaseSelected + "SegmentoMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de segmentos MT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLine("SMT_" + sqlDtRdrServer["CodSegmMT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["De"].ToString(), sqlDtRdrServer["Para"].ToString(), sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(sqlDtRdrServer["Comp_km"]), sqlDtRdrServer["Propr"].ToString(), 0, intNeutralizarRedeTerceiros);
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos segmentos MT.");

                        // Loop dos reguladores MT
                        sqlCommServer = new SqlCommand("SELECT t1.CodRegulMT, t1.CodBnc, t1.TipRegul, t1.CodFasPrim, t1.CodFasSecu, t1.PotNom_kVA, t1.TenRgl_pu, t1.[ReatHL_%], t1.PerdTtl_W, t1.PerdVz_W, t1.TnsLnh1_kV, t1.De, t1.Para, t1.CodFasCoinc FROM dbo.[" + strBaseSelected + "ReguladorMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de reguladores MT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandRegulatorMT("REG_" + sqlDtRdrServer["CodRegulMT"].ToString(), Convert.ToInt32(sqlDtRdrServer["CodBnc"]), Convert.ToInt32(sqlDtRdrServer["TipRegul"]), sqlDtRdrServer["CodFasPrim"].ToString(), sqlDtRdrServer["CodFasSecu"].ToString(), sqlDtRdrServer["CodFasCoinc"].ToString(), sqlDtRdrServer["De"].ToString(), sqlDtRdrServer["Para"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"]), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(sqlDtRdrServer["TenRgl_pu"]), Convert.ToDouble(sqlDtRdrServer["ReatHL_%"]), 100 * Convert.ToDouble(sqlDtRdrServer["PerdTtl_W"]) / (1000 * Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"])), 100 * Convert.ToDouble(sqlDtRdrServer["PerdVz_W"]) / (1000 * Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"])));
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos reguladores MT.");

                        // Loop dos transformadores MT-MT ou MT-BT
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + strBaseSelected + "TrafoMTMTMTBT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de transformadores MT-MT ou MT-BT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandTransformer("TRF_" + sqlDtRdrServer["CodTrafo"].ToString(), Convert.ToInt32(sqlDtRdrServer["CodBnc"]), Convert.ToInt32(sqlDtRdrServer["MRT"]), Convert.ToInt32(sqlDtRdrServer["TipTrafo"]), sqlDtRdrServer["CodFasPrim"].ToString(), sqlDtRdrServer["CodFasSecu"].ToString(), sqlDtRdrServer["CodFasTerc"].ToString(), sqlDtRdrServer["De"].ToString(), sqlDtRdrServer["Para"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"]), Convert.ToDouble(sqlDtRdrServer["TenSecu_kV"]), Convert.ToDouble(sqlDtRdrServer["Tap_pu"]), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), GeoPerdasTools.mthdTransformerTotalPowerLoss_per(sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"])) != 0 && intUsaTrafoABNT == 1 ? GeoPerdasTools.mthdTransformerTotalPowerLoss_per(sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"])) : Convert.ToDouble(sqlDtRdrServer["PerdTtl_%"]), GeoPerdasTools.mthdTransformerNoLoadPowerLoss_per(sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"])) != 0 && intUsaTrafoABNT == 1 ? GeoPerdasTools.mthdTransformerNoLoadPowerLoss_per(sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh1_kV"])) : Convert.ToDouble(sqlDtRdrServer["PerdVz_%"]), sqlDtRdrServer["Propr"].ToString(), intAdequarTrafoVazio == 1 ? Convert.ToInt32(sqlDtRdrServer["SemCarga"]) : 0);
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos transformadores MT-MT ou MT-BT.");

                        // Loop das chaves BT
                        sqlCommServer = new SqlCommand("SELECT t1.CodChvBT, t1.CodFas, t1.De, t1.Para FROM dbo.[" + strBaseSelected + "ChaveBT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "' AND t1.EstChv = 2", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de chaves BT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandMTSwitch("CBT_" + sqlDtRdrServer["CodChvBT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["De"].ToString(), sqlDtRdrServer["Para"].ToString());
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelas chaves BT.");

                        // Loop dos segmentos BT
                        sqlCommServer = new SqlCommand("SELECT t1.*, [dbo].[EscolhaPropriedade]([t2].[M1_Propr], [t2].[M2_Propr], [t2].[M3_Propr]) AS [Propr] FROM dbo.[" + strBaseSelected + "SegmentoBT] AS t1 LEFT JOIN [dbo].[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON ([t1].[CodBase] = [t2].[M1_CodBase] AND [t1].[CodTrafoAtrib] = [t2].[M1_CodTrafo]) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de segmentos BT");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLine("SBT_" + sqlDtRdrServer["CodSegmBT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl1"].ToString(), sqlDtRdrServer["CodPonAcopl2"].ToString(), sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(sqlDtRdrServer["Comp_km"]), sqlDtRdrServer["Propr"].ToString(), 0, intNeutralizarRedeTerceiros);
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos segmentos BT.");

                        // Loop dos ramais
                        sqlCommServer = new SqlCommand("SELECT t1.*, [dbo].[EscolhaPropriedade]([t2].[M1_Propr], [t2].[M2_Propr], [t2].[M3_Propr]) AS [Propr] FROM dbo.[" + strBaseSelected + "RamalBT] AS t1 LEFT JOIN [dbo].[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON ([t1].[CodBase] = [t2].[M1_CodBase] AND [t1].[CodTrafoAtrib] = [t2].[M1_CodTrafo]) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de ramais");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            this.mthdGetDSSCommandLine("RBT_" + sqlDtRdrServer["CodRmlBT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl1"].ToString(), sqlDtRdrServer["CodPonAcopl2"].ToString(), sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(sqlDtRdrServer["Comp_km"]), sqlDtRdrServer["Propr"].ToString(), intAdequarRamal, intNeutralizarRedeTerceiros);
                        }
                        sqlDtRdrServer.Close();
                        mthdInputMensageLog("Passou pelos ramais.");

                        // Versão usando apenas o medidor de barramento
                        if (config.cbMeterComplete == false)
                        {
                            // Loop dos energymeters de barramento
                            i = 0;
                            sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxMeterBarra AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                            sqlCommServer.CommandTimeout = 0;
                            mthdInsertCommandInLine("! Criação da seção de energymeters de barramento");
                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                            while (sqlDtRdrServer.Read())
                            {
                                this.mthdGetDSSCommandEnergymeter(sqlDtRdrServer["Nome"].ToString().TrimEnd(), this.mthdNomeElemento(sqlDtRdrServer["Elem"].ToString().TrimEnd()) + sqlDtRdrServer["CodNomeElem"].ToString().TrimEnd());
                                i++;
                            }
                            sqlDtRdrServer.Close();
                            mthdInputMensageLog("Passou pelos energymeters de barramento.");
                        }
                        // Versão usando a tabela auxiliar de medidores
                        if (config.cbMeterComplete == true)
                        {
                            // Loop dos energymeters usando AuxMeter
                            sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxMeterCompleto AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                            sqlCommServer.CommandTimeout = 0;
                            mthdInsertCommandInLine("! Criação da seção de energymeters");
                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                            while (sqlDtRdrServer.Read())
                            {
                                this.mthdGetDSSCommandEnergymeter(sqlDtRdrServer["Nome"].ToString().TrimEnd(), this.mthdNomeElemento(sqlDtRdrServer["Elem"].ToString().TrimEnd()) + sqlDtRdrServer["CodNomeElem"].ToString().TrimEnd());
                            }
                            sqlDtRdrServer.Close();
                            mthdInputMensageLog("Passou pelos energymeters.");
                        }
                        // Loop das tensões de base
                        strTensoesBase = "Set voltagebases=[";
                        i = 0;
                        sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxVoltageBases AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                        sqlCommServer.CommandTimeout = 0;
                        mthdInsertCommandInLine("! Criação da seção de tensões de base");
                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                        while (sqlDtRdrServer.Read())
                        {
                            if (i == 0)
                                strTensoesBase += sqlDtRdrServer["TnsLnh_kV"].ToString().Replace(",", ".");
                            else
                                strTensoesBase += " " + sqlDtRdrServer["TnsLnh_kV"].ToString().Replace(",", ".");
                            i++;
                        }
                        sqlDtRdrServer.Close();
                        strTensoesBase += "]";
                        mthdInsertCommandInLine(strTensoesBase);
                        mthdInsertCommandInLine("Calcvoltagebases");
                        mthdInputMensageLog("Passou pelas tensões de base.");

                    ModeloConvergenciaSecundaria:

                        if (intModeloConverge == 1)
                        {
                            intCtrlCnvrg = 0;
                        }
                        else if (intModeloConverge == 2)
                        {
                            // Zera as cargas não técnicas calculadas anteriormente para reiniciar o processo
                            if (config.cbReinicializeNT == true)
                            {
                                sqlCommServer = new SqlCommand("dbo.iadAuxCargaBTNTDem", sqlConnServer);
                                sqlCommServer.CommandType = CommandType.StoredProcedure;
                                sqlCommServer.Parameters.AddWithValue("@CodBase", strBaseSelected);
                                sqlCommServer.Parameters.AddWithValue("@CodOperation", 1);
                                sqlCommServer.Parameters.AddWithValue("@CodAlim", strFeeder);
                                sqlCommServer.ExecuteNonQuery();

                                sqlCommServer = new SqlCommand("dbo.iadAuxCargaMTNTDem", sqlConnServer);
                                sqlCommServer.CommandType = CommandType.StoredProcedure;
                                sqlCommServer.Parameters.AddWithValue("@CodBase", strBaseSelected);
                                sqlCommServer.Parameters.AddWithValue("@CodOperation", 1);
                                sqlCommServer.Parameters.AddWithValue("@CodAlim", strFeeder);
                                sqlCommServer.ExecuteNonQuery();
                            }

                            intCtrlCnvrg = 1;

                            // Variáveis de avaliação de troca de modelo
                            dblEnergyInjectDeclaredYearly_MWh = 0;
                            dblEnergyInjectCalculatedYearly_MWh = 0;
                        }

                        for (int j = 1; j <= 12; j++)
                        {
                            // Inicia o Loop da convergência não técnica
                            intQuantIntrtn = 0;
                            do
                            {
                                // Captura a energia injetada declarada pelo agente
                                sqlCommServer = new SqlCommand("SELECT t1.EnerCirc" + string.Format("{0:00}", j) + "_MWh FROM dbo.[" + strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                                sqlCommServer.CommandTimeout = 0;
                                sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                sqlDtRdrServer.Read();
                                dblEnergyInjectDeclared_MWh = Convert.ToDouble(sqlDtRdrServer["EnerCirc" + string.Format("{0:00}", j) + "_MWh"]);
                                sqlDtRdrServer.Close();
                                mthdInputMensageLog("Passou pelas obtenção da energia injetada declarada.");

                                // Captura a energia das cargas MT declaradas pelo agente
                                sqlCommServer = new SqlCommand("SELECT Sum(t1.EnerMedid" + string.Format("{0:00}", j) + "_MWh) AS EnerMedid FROM dbo.[" + strBaseSelected + "CargaMT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                sqlCommServer.CommandTimeout = 0;
                                sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                sqlDtRdrServer.Read();
                                dblEnergyLoadsMTDeclared_MWh = DBNull.Value.Equals(sqlDtRdrServer["EnerMedid"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["EnerMedid"]);
                                sqlDtRdrServer.Close();
                                mthdInputMensageLog("Passou pelas obtenção da energia cargas MT declaradas.");

                                // Captura a energia das cargas BT declaradas pelo agente
                                sqlCommServer = new SqlCommand("SELECT Sum(t1.EnerMedid" + string.Format("{0:00}", j) + "_MWh) AS EnerMedid FROM dbo.[" + strBaseSelected + "CargaBT] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                sqlCommServer.CommandTimeout = 0;
                                sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                sqlDtRdrServer.Read();
                                dblEnergyLoadsBTDeclared_MWh = DBNull.Value.Equals(sqlDtRdrServer["EnerMedid"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["EnerMedid"]);
                                sqlDtRdrServer.Close();
                                mthdInputMensageLog("Passou pelas obtenção da energia cargas BT declaradas.");

                                if (intRealizaCnvrgcPNT == 1)
                                {
                                    // Captura a energia das cargas MT Não Técnica declaradas pelo agente
                                    sqlCommServer = new SqlCommand("SELECT Sum(t1.EnerMedid" + string.Format("{0:00}", j) + "_MWh) AS EnerMedid FROM dbo.[" + strBaseSelected + "AuxCargaMTNTDem" + strFeeder + "] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                    sqlDtRdrServer.Read();
                                    dblEnergyLoadsMTNTCalculated_MWh = DBNull.Value.Equals(sqlDtRdrServer["EnerMedid"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["EnerMedid"]);
                                    sqlDtRdrServer.Close();
                                    mthdInputMensageLog("Passou pelas obtenção da energia cargas MT Não Técnicas declaradas.");

                                    // Captura a energia das cargas BT Não Técnica declaradas pelo agente
                                    sqlCommServer = new SqlCommand("SELECT Sum(t1.EnerMedid" + string.Format("{0:00}", j) + "_MWh) AS EnerMedid FROM dbo.[" + strBaseSelected + "AuxCargaBTNTDem" + strFeeder + "] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                    sqlDtRdrServer.Read();
                                    dblEnergyLoadsBTNTCalculated_MWh = DBNull.Value.Equals(sqlDtRdrServer["EnerMedid"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["EnerMedid"]);
                                    sqlDtRdrServer.Close();
                                    mthdInputMensageLog("Passou pelas obtenção da energia cargas BT Não Técnicas declaradas.");

                                    // Captura a proporção da energia que deve ficar no MT
                                    sqlCommServer = new SqlCommand("SELECT t1.PropPerdNTecnMT" + string.Format("{0:00}", j) + "_pu AS PropPerdNTecnMT FROM dbo.AuxProporPerdNTecn AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                    sqlDtRdrServer.Read();
                                    dblPropEnergyLoadsMT_pu = DBNull.Value.Equals(sqlDtRdrServer["PropPerdNTecnMT"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["PropPerdNTecnMT"]);
                                    sqlDtRdrServer.Close();
                                    mthdInputMensageLog("Passou pelas obtenção da proporção de energia não técnica de MT.");

                                    // Captura a proporção da energia que deve ficar no BT
                                    sqlCommServer = new SqlCommand("SELECT t1.PropPerdNTecnBT" + string.Format("{0:00}", j) + "_pu AS PropPerdNTecnBT FROM dbo.AuxProporPerdNTecn AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                    sqlDtRdrServer.Read();
                                    dblPropEnergyLoadsBT_pu = DBNull.Value.Equals(sqlDtRdrServer["PropPerdNTecnBT"]) ? 0 : Convert.ToDouble(sqlDtRdrServer["PropPerdNTecnBT"]);
                                    sqlDtRdrServer.Close();
                                    mthdInputMensageLog("Passou pelas obtenção da proporção de energia não técnica de BT.");
                                }

                                for (int k = 1; k <= 3; k++)
                                {
                                    if (j == 1 && k == 1 && intCtrlCnvrg == 0)
                                    {
                                        // Flag que sinaliza que a carga não deve ser escrita novamente quando houver necessidade de convergência da NT
                                        intCtrlCnvrg = 1;

                                        // Loop das cargas MT
                                        sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaMTDem] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                        sqlCommServer.CommandTimeout = 0;
                                        mthdInsertCommandInLine("! Criação da seção de cargas MT");
                                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                        while (sqlDtRdrServer.Read())
                                        {
                                            this.mthdGetDSSCommandNewLoad("MT_" + sqlDtRdrServer["CodConsMT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, intAdequarTensaoCargasMT, 0);
                                        }
                                        sqlDtRdrServer.Close();
                                        mthdInputMensageLog("Passou pelas cargas MT.");

                                        // Loop das cargas BT
                                        sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV, t2.TnsFasBas_kV FROM dbo.[" + strBaseSelected + "AuxCargaBTDem] AS t1 LEFT JOIN dbo.[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                        sqlCommServer.CommandTimeout = 0;
                                        mthdInsertCommandInLine("! Criação da seção de cargas BT");
                                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                        while (sqlDtRdrServer.Read())
                                        {
                                            this.mthdGetDSSCommandNewLoad("BT_" + sqlDtRdrServer["CodConsBT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["TnsFasBas_kV"]), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(sqlDtRdrServer["DemMaxTrafo_kW"]), sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, intAdequarTensaoCargasBT, 0);
                                        }
                                        sqlDtRdrServer.Close();
                                        mthdInputMensageLog("Passou pelas cargas BT.");

                                        if (intRealizaCnvrgcPNT == 1)
                                        {
                                            // Loop das cargas MT Não Técnicas
                                            sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaMTNTDem" + strFeeder + "] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                            sqlCommServer.CommandTimeout = 0;
                                            mthdInsertCommandInLine("! Criação da seção de cargas MT Não Técnicas");
                                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                            while (sqlDtRdrServer.Read())
                                            {
                                                this.mthdGetDSSCommandNewLoad("MT_" + sqlDtRdrServer["CodConsMT"].ToString() + "_NT", sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, intAdequarTensaoCargasMT, 0);
                                            }
                                            sqlDtRdrServer.Close();
                                            mthdInputMensageLog("Passou pelas cargas MT Não Técnicas.");

                                            // Loop das cargas BT Não Técnicas
                                            sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV, t2.TnsFasBas_kV FROM dbo.[" + strBaseSelected + "AuxCargaBTNTDem" + strFeeder + "] AS t1 LEFT JOIN dbo.[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                            sqlCommServer.CommandTimeout = 0;
                                            mthdInsertCommandInLine("! Criação da seção de cargas BT Não Técnicas");
                                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                            while (sqlDtRdrServer.Read())
                                            {
                                                this.mthdGetDSSCommandNewLoad("BT_" + sqlDtRdrServer["CodConsBT"].ToString() + "_NT", sqlDtRdrServer["CodFas"].ToString(), sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(sqlDtRdrServer["TnsFasBas_kV"]), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(sqlDtRdrServer["DemMaxTrafo_kW"]), sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, intAdequarTensaoCargasBT, 0);
                                            }
                                            sqlDtRdrServer.Close();
                                            mthdInputMensageLog("Passou pelas cargas BT Não Técnicas.");
                                        }
                                    }
                                    else
                                    {
                                        // Reinício do código
                                        mthdStarCommands();

                                        // Loop das cargas MT
                                        sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaMTDem] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                        sqlCommServer.CommandTimeout = 0;
                                        mthdInsertCommandInLine("! Criação da seção de cargas MT");
                                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                        while (sqlDtRdrServer.Read())
                                        {
                                            this.mthdGetDSSCommandEditLoad("MT_" + sqlDtRdrServer["CodConsMT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, 0);
                                        }
                                        sqlDtRdrServer.Close();
                                        mthdInputMensageLog("Passou pelas cargas MT.");

                                        // Loop das cargas BT
                                        sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaBTDem] AS t1 LEFT JOIN dbo.[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                        sqlCommServer.CommandTimeout = 0;
                                        mthdInsertCommandInLine("! Criação da seção de cargas BT");
                                        sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                        while (sqlDtRdrServer.Read())
                                        {
                                            this.mthdGetDSSCommandEditLoad("BT_" + sqlDtRdrServer["CodConsBT"].ToString(), sqlDtRdrServer["CodFas"].ToString(), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(sqlDtRdrServer["DemMaxTrafo_kW"]), sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, 0);
                                        }
                                        sqlDtRdrServer.Close();
                                        mthdInputMensageLog("Passou pelas cargas BT.");

                                        if (intRealizaCnvrgcPNT == 1)
                                        {
                                            // Loop das cargas MT Não Técnicas
                                            sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaMTNTDem" + strFeeder + "] AS t1 WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                            sqlCommServer.CommandTimeout = 0;
                                            mthdInsertCommandInLine("! Criação da seção de cargas MT Não Técnicas");
                                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                            while (sqlDtRdrServer.Read())
                                            {
                                                this.mthdGetDSSCommandEditLoad("MT_" + sqlDtRdrServer["CodConsMT"].ToString() + "_NT", sqlDtRdrServer["CodFas"].ToString(), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, 0);
                                            }
                                            sqlDtRdrServer.Close();
                                            mthdInputMensageLog("Passou pelas cargas MT Não Técnicas.");

                                            // Loop das cargas BT Não Técnicas
                                            sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + strBaseSelected + "AuxCargaBTNTDem" + strFeeder + "] AS t1 LEFT JOIN dbo.[" + strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", sqlConnServer);
                                            sqlCommServer.CommandTimeout = 0;
                                            mthdInsertCommandInLine("! Criação da seção de cargas BT Não Técnicas");
                                            sqlDtRdrServer = sqlCommServer.ExecuteReader();
                                            while (sqlDtRdrServer.Read())
                                            {
                                                this.mthdGetDSSCommandEditLoad("BT_" + sqlDtRdrServer["CodConsBT"].ToString() + "_NT", sqlDtRdrServer["CodFas"].ToString(), Convert.ToDouble(sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(sqlDtRdrServer["DemMaxTrafo_kW"]), sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, 0);
                                            }
                                            sqlDtRdrServer.Close();
                                            mthdInputMensageLog("Passou pelas cargas BT Não Técnicas.");
                                        }
                                    }

                                    // Ajustes de execução da apuração
                                    tolerance = 0.00001;
                                    do
                                    {
                                        tolerance = tolerance * 10;
                                        mthdInsertCommandInLine("Set mode = daily");
                                        mthdInsertCommandInLine("Set tolerance = " + tolerance.ToString().Replace(",", "."));
                                        mthdInsertCommandInLine("Set maxcontroliter = 10");
                                        mthdInsertCommandInLine("!Set algorithm = newton");
                                        mthdInsertCommandInLine("!Solve mode = direct");
                                        mthdInsertCommandInLine("Solve");
                                        mthdEndCommands();
                                    } while (intConverged == 0);

                                    // Caputa as energia do dia útil, sábado e domingo para a avaliação da convergência não técnica
                                    if (k == 1)
                                    {
                                        dblEnerInjDU_kWh = dblTotalEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblPerdaEnerTecDU_kWh = dblTecnicalLossEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblEnerForncDU_kWh = (dblMTLoad_kWh + dblBTLoad_kWh) * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                    }
                                    else if (k == 2)
                                    {
                                        dblEnerInjSA_kWh = dblTotalEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblPerdaEnerTecSA_kWh = dblTecnicalLossEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblEnerForncSA_kWh = (dblMTLoad_kWh + dblBTLoad_kWh) * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                    }
                                    else if (k == 3)
                                    {
                                        dblEnerInjDO_kWh = dblTotalEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblPerdaEnerTecDO_kWh = dblTecnicalLossEnergy_kWh * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                        dblEnerForncDO_kWh = (dblMTLoad_kWh + dblBTLoad_kWh) * GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k));
                                    }

                                    // Armazena os resultados na tabela do servidor
                                    sqlCommServer = new SqlCommand("DELETE FROM dbo.[" + strBaseSelected + "AuxResultado] WHERE CodBase = " + strBaseSelected + " AND CodAlim = '" + strFeeder + "' AND TipoDiaMes = '" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "' AND TipoRodada = '" + mthdGetOptionsToString() + "'", sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlCommServer.ExecuteNonQuery();
                                    strTempSqlCommand = "INSERT INTO dbo.[" + strBaseSelected + "AuxResultado] (CodBase, CodAlim, TipoDiaMes, TipoRodada, Dias, Tolerancia, PotenciaMaxInj_kW, EnergiaInj_kWh, EnergiaFornc_kWh, PerdaPotenciaMaxTecnica_kW, PerdaEnergiaTecnica_kWh, PerdaEnergiaFerro_kWh, PerdaEnergiaTrafo_kWh, PerdaEnergiaFerroA3aA3a_kWh, PerdaEnergiaTrafoA3aA3a_kWh, PerdaEnergiaFerroA3aA4_kWh, PerdaEnergiaTrafoA3aA4_kWh, PerdaEnergiaFerroA3aB_kWh, PerdaEnergiaTrafoA3aB_kWh, PerdaEnergiaFerroA4A4_kWh, PerdaEnergiaTrafoA4A4_kWh, PerdaEnergiaFerroA4B_kWh, PerdaEnergiaTrafoA4B_kWh, PerdaEnergiaFerroA4A3a_kWh, PerdaEnergiaTrafoA4A3a_kWh, PerdaEnergiaFerroBA3a_kWh, PerdaEnergiaTrafoBA3a_kWh, PerdaEnergiaFerroBA4_kWh, PerdaEnergiaTrafoBA4_kWh, PerdaEnergiaLinhas_kWh, PerdaEnergiaMT_kWh, PerdaEnergiaBT_kWh, PerdaEnergiaLinhasMT_kWh, PerdaEnergiaLinhasBT_kWh, PerdaEnergiaLinhasMTA3a_kWh, PerdaEnergiaLinhasMTA4_kWh, EnergiaForncMT_kWh, EnergiaForncBT_kWh) ";
                                    strTempSqlCommand += "VALUES (" + strBaseSelected + ", '" + strFeeder + "', '" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "', '" + mthdGetOptionsToString() + "', " + GeoPerdasTools.mthdGetNumberDaysDUSADOMonth(j, intAno, GeoPerdasTools.mthdGetTypeDay(k)) + ", " + tolerance.ToString().Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTotalPower_kW)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTotalEnergy_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTotalLoadEnergy_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTecnicalLossPower_kW)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTecnicalLossEnergy_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aA3aTransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aA3aTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aA4TransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aA4TransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aBTransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA3aBTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4A4TransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4A4TransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4BTransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4BTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4A3aTransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblA4A3aTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBA3aTransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBA3aTransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBA4TransformerNoLoadLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBA4TransformerLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblLineLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblMTLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBTLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblLineMTLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblLineBTLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblLineA3aLosses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblLineA4Losses_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblMTLoad_kWh)).Replace(",", ".") + ", " + string.Format("{0:F9}", GeoPerdasTools.mthdNumberCorrection(dblBTLoad_kWh)).Replace(",", ".") + ")";
                                    sqlCommServer = new SqlCommand(strTempSqlCommand, sqlConnServer);
                                    sqlCommServer.CommandTimeout = 0;
                                    sqlCommServer.ExecuteNonQuery();
                                }

                                mthdInputMensageLog(GeoPerdasTools.mthdNumberCorrection(dblEnergyInjectDeclared_MWh).ToString() + " = " + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh).ToString() + "+" + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh).ToString() + "+" + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh).ToString() + "+" + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh).ToString() + "+" + (GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecDU_kWh) / 1000 + GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecSA_kWh) / 1000 + GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecDO_kWh) / 1000).ToString());

                                if (intModeloConverge == 1)
                                {
                                    dblDelta = (GeoPerdasTools.mthdNumberCorrection(dblEnergyInjectDeclared_MWh) - GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) - GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh) - GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh) - GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh) - GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecDU_kWh) / 1000 - GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecSA_kWh) / 1000 - GeoPerdasTools.mthdNumberCorrection(dblPerdaEnerTecDO_kWh) / 1000) / (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh));
                                }
                                else if (intModeloConverge == 2)
                                {
                                    dblDelta = (GeoPerdasTools.mthdNumberCorrection(dblEnergyInjectDeclared_MWh) - dblEnerInjDU_kWh / 1000 - dblEnerInjSA_kWh / 1000 - dblEnerInjDO_kWh / 1000) / (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh));

                                    //Dec alim - (EneInjCalc) / DecCargas + PNTInj
                                }
                                dblDelta_MWh = dblDelta * (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh));
                                dblDeltaMT_MWh = dblDelta_MWh * dblPropEnergyLoadsMT_pu;
                                dblDeltaBT_MWh = dblDelta_MWh * dblPropEnergyLoadsBT_pu;

                                if (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh) != 0)
                                    dblDeltaMT = dblDeltaMT_MWh / (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsMTNTCalculated_MWh));
                                else
                                    dblDeltaMT = 0;

                                if (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh) != 0)
                                    dblDeltaBT = dblDeltaBT_MWh / (GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTDeclared_MWh) + GeoPerdasTools.mthdNumberCorrection(dblEnergyLoadsBTNTCalculated_MWh));
                                else
                                    dblDeltaBT = 0;

                                mthdInputMensageLog("Delta (MWh): " + dblDelta_MWh.ToString());

                                if (intRealizaCnvrgcPNT == 1)
                                {
                                    // Executa a stored procedure de atualização das cargas BT Não Técnicas
                                    sqlCommServer = new SqlCommand("dbo.AtualizaDemandaCargaBTNT", sqlConnServer);
                                    sqlCommServer.CommandType = CommandType.StoredProcedure;
                                    sqlCommServer.Parameters.AddWithValue("@BaseAnalisada", strBaseSelected);
                                    sqlCommServer.Parameters.AddWithValue("@AlimAnalisado", strFeeder);
                                    sqlCommServer.Parameters.AddWithValue("@Mes", j);
                                    sqlCommServer.Parameters.AddWithValue("@DeltaBT", dblDeltaBT);
                                    sqlCommServer.ExecuteNonQuery();
                                    sqlCommServer.ExecuteNonQuery();

                                    // Executa a stored procedure de atualização das cargas MT Não Técnicas
                                    sqlCommServer = new SqlCommand("dbo.AtualizaDemandaCargaMTNT", sqlConnServer);
                                    sqlCommServer.CommandType = CommandType.StoredProcedure;
                                    sqlCommServer.Parameters.AddWithValue("@BaseAnalisada", strBaseSelected);
                                    sqlCommServer.Parameters.AddWithValue("@AlimAnalisado", strFeeder);
                                    sqlCommServer.Parameters.AddWithValue("@Mes", j);
                                    sqlCommServer.Parameters.AddWithValue("@DeltaMT", dblDeltaMT);
                                    sqlCommServer.ExecuteNonQuery();
                                }

                                intQuantIntrtn++;
                            } while (intRealizaCnvrgcPNT == 1 && Math.Abs(dblDelta_MWh) > 0.5 && intQuantIntrtn <= 10);

                            dblEnergyInjectDeclaredYearly_MWh += GeoPerdasTools.mthdNumberCorrection(dblEnergyInjectDeclared_MWh);
                            dblEnergyInjectCalculatedYearly_MWh += (dblEnerInjDU_kWh + dblEnerInjSA_kWh + dblEnerInjDO_kWh) / 1000;

                            if (dblDelta_MWh > 0.5)
                                mthdInputMensageLog("O processo realizou " + intQuantIntrtn.ToString() + " iterações e foi concluído foi concluído com energia superior 0.5 Mwh: " + dblDelta_MWh.ToString());
                            else
                                mthdInputMensageLog("O processo realizou " + intQuantIntrtn.ToString() + " iterações e foi concluído foi concluído com energia inferior ou igual 0.5 Mwh: " + dblDelta_MWh.ToString());
                        }

                        dblErrorEnerInj = (dblEnergyInjectDeclaredYearly_MWh - dblEnergyInjectCalculatedYearly_MWh) / dblEnergyInjectDeclaredYearly_MWh;
                        if (intRealizaCnvrgcPNT == 1 && intModeloConverge == 1 && dblErrorEnerInj < 7.3 / 100)
                        {
                            mthdInputMensageLog("O processo precisa ser executado com modelo de convergência secundário (erro " + (dblErrorEnerInj * 100).ToString() + "% inferior a 7.3%)");
                            intModeloConverge = 2;
                            goto ModeloConvergenciaSecundaria;
                        }

                        // Executa a stored procedure de atualização dos resultados mensais e anuais do alimentador
                        sqlCommServer = new SqlCommand("dbo.AtualizaResultadosAlimentador", sqlConnServer);
                        sqlCommServer.CommandType = CommandType.StoredProcedure;
                        sqlCommServer.Parameters.AddWithValue("@BaseAnalisada", strBaseSelected);
                        sqlCommServer.Parameters.AddWithValue("@AlimAnalisado", strFeeder);
                        sqlCommServer.Parameters.AddWithValue("@TipoRodada", mthdGetOptionsToString());
                        sqlCommServer.ExecuteNonQuery();

                        if (intRealizaCnvrgcPNT == 1)
                        {
                            // Executa a stored procedure de armazenamento final dos resultados das cargas nao tecnicas do alimentador
                            sqlCommServer = new SqlCommand("dbo.ArmazenaCargaNTAlimentador", sqlConnServer);
                            sqlCommServer.CommandType = CommandType.StoredProcedure;
                            sqlCommServer.Parameters.AddWithValue("@BaseAnalisada", strBaseSelected);
                            sqlCommServer.Parameters.AddWithValue("@AlimAnalisado", strFeeder);
                            sqlCommServer.ExecuteNonQuery();
                        }

                        // Término do processo
                        mthdInputMensageLog("Término do processo para o alimentador " + strFeeder + ".");
                    }
                    catch (Exception xcptn)
                    {
                        mthdInputMensageLog("Problema não identificado no alimentador " + strFeeder + " (" + xcptn.Message.ToString() + ").");
                        sqlDtRdrServer.Dispose();
                    }

                    /*
                    this.strFileData[this.lngindex] = "<FIM DOS DADOS>"; this.lngindex++;
                    using (swFile = new System.IO.StreamWriter(@"C:\Users\renatosousa\Desktop\Teste_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".txt", false, Encoding.GetEncoding(1252)))
                    {
                        foreach (String strLine in this.strFileData)
                        {
                            if (strLine != "<FIM DOS DADOS>")
                                swFile.WriteLine(strLine);
                            else
                                break;
                        }
                    }
                    */
                }
          
            }
            catch (Exception xcptn)
            {
                mthdInputMensageLog("Problema não identificado (" + xcptn.Message.ToString() + ").");
                sqlDtRdrServer.Dispose();
                sqlCommServer.Dispose();
            }
        }
    }

    private void mthdOpenDataBase(string strServer, string strDataBase, string usuario, string senha)
    {
        try
        {
            strCommandConnection = "Data Source=" + strServer;
            strCommandConnection += ";Initial Catalog=" + strDataBase;
            strCommandConnection += $";User Id={usuario};Password={senha}";
            sqlConnServer = new SqlConnection(strCommandConnection);

            sqlConnServer.Open();
            mthdInputMensageLog("A conexão com o banco de dados foi estabelecida.");
        }
        catch (Exception xcptn)
        {
            mthdInputMensageLog("Não foi possível estabelecer conexão com o banco de dados (" + xcptn.Message.ToString() + ").");
        }
    }

    private void mthdCloseDataBase()
    {
        if (sqlConnServer != null)
            sqlConnServer.Close();
        else
            mthdInputMensageLog("Não há qualquer conexão ativa.");
    }

    private void mthdInputMensageLog(string message)
    {
        _logger.LogInformation("[" + DateTime.Now.ToString() + "] " + message);
    }

    private void mthdStarDSSServer()
    {
        _dss = new DssWrapper();
        _dss.InitiateServer();
        if (!_dss.DSSServer.Start(0))
        {
            mthdInputMensageLog("A inicialização do servidor DSS falhou.");
            return;
        }
        else
        {
            mthdInputMensageLog("A inicialização do servidor DSS obteve sucesso.");
            _dss.Text = _dss.DSSServer.Text;
            _dss.DSSServer.AllowForms = false;
        }
    }

    private void mthdLoadCircuit(string strPathFile)
    {
        if (strPathFile == "" || strPathFile == null)
        {
            mthdInputMensageLog("O caminho e o arquivo não foi informado.");
            return;
        }
        if (!_dss.DSSServer.Start(0))
        {
            mthdInputMensageLog("O servidor DSS não foi inicializado.");
            return;
        }

        _dss.Text.Command = "clear";
        _dss.Text.Command = "compile " + "(" + strPathFile + ")";
        _dss.Circuit = _dss.DSSServer.ActiveCircuit;
        _dss.Solution = _dss.Circuit.Solution;
        _dss.SolutionControlQueue = _dss.Circuit.CtrlQueue;

        mthdInputMensageLog("O circuito foi carregado.");
    }

    private void mthdStarCommands()
    {
        if (!_dss.DSSServer.Start(0))
        {
            mthdInputMensageLog("O servidor DSS não foi inicializado.");
            return;
        }
    }

    private void mthdInsertCommandInLine(string strCommand)
    {
        if (strCommand == "" || strCommand == null)
        {
            mthdInputMensageLog("O comando não foi informado.");
            return;
        }
        if (!_dss.DSSServer.Start(0))
        {
            mthdInputMensageLog("O servidor DSS não foi inicializado.");
            return;
        }
        //this.strFileData[this.lngindex] = strCommand; this.lngindex++;
        _dss.Text.Command = strCommand;
    }

    private void mthdEndCommands()
    {
        if (!_dss.DSSServer.Start(0))
        {
            mthdInputMensageLog("O servidor DSS não foi inicializado.");
            return;
        }
        _dss.Circuit = _dss.DSSServer.ActiveCircuit;
        _dss.Solution = _dss.Circuit.Solution;
        _dss.SolutionControlQueue = _dss.Circuit.CtrlQueue;
        if (_dss.Solution.Converged)
        {
            intConverged = 1;
            mthdInputMensageLog("Solução convergiu.");
            //this.mthdInputMensageLog("\t" + "Tensões de Sequencia");
            //this.mthdInputMensageLog("\t" + "Barra" + "\t" + "Módulo da Tensão de Sequencia 0 (kV)" + "\t" + "Módulo da Tensão de Sequencia Positiva (kV)" + "\t" + "Módulo da Tensão de Sequencia Negativa (kV)" + "\t" + "V0.Real (kV)" + "\t" + "V0.Imaginário (kV)" + "\t" + "V1.Real (kV)" + "\t" + "V1.Imaginário (kV)" + "\t" + "V2.Real (kV)" + "\t" + "V2.Imaginário (kV)");
            //this.getSequenceVoltages();
            //this.mthdInputMensageLog("\t" + "Tensões Complexas por Fase (pu)");
            //this.mthdInputMensageLog("\t" + "Barra" + "\t" + "(Tensão Fase 1).Real (pu)" + "\t" + "(Tensão Fase 1).Imaginária (pu)" + "\t" + "(Tensão Fase 2).Real (pu)" + "\t" + "(Tensão Fase 2).Imaginária (pu)" + "\t" + "(Tensão Fase 3).Real (pu)" + "\t" + "(Tensão Fase 3).Imaginária (pu)");
            //this.getVoltages();
            //this.mthdInputMensageLog("\t" + "Tensões Polares por Fase (pu)");
            //this.mthdInputMensageLog("\t" + "Barra" + "\t" + "(Módulo da Tensão da Fase 1 ; Ângulo da Tensão da Fase 1) (pu)" + "\t" + "(Módulo da Tensão da Fase 2 ; Ângulo da Tensão da Fase 2) (pu)" + "\t" + "(Módulo da Tensão da Fase 3 ; Ângulo da Tensão da Fase 3) (pu)");
            //this.getMagnitudeAngle();
            //this.mthdInputMensageLog("\t" + "Correntes de Sequencia");
            //this.getSequenceCurrents();
            //this.mthdInputMensageLog("\t" + "Potências");
            //this.mthdInputMensageLog("\t" + "Carga / Geração" + "\t" + "(Potência Ativa (kW) ; Potência Reativa (kvar))");
            //this.getLoadPowers();
            mthdGetInfoMeters();
            mthdInputMensageLog("\tPotenciaMaxInj_kW\tEnergiaInj_kWh\tEnergiaFornc_kWh\tPerdaPotenciaMaxTecnica_kW\tPerdaEnergiaTecnica_kWh\tPerdaEnergiaFerro_kWh\tPerdaEnergiaTrafo_kWh\tPerdaEnergiaFerroA3aA3a_kWh\tPerdaEnergiaTrafoA3aA3a_kWh\tPerdaEnergiaFerroA3aA4_kWh\tPerdaEnergiaTrafoA3aA4_kWh\tPerdaEnergiaFerroA3aB_kWh\tPerdaEnergiaTrafoA3aB_kWh\tPerdaEnergiaFerroA4A4_kWh\tPerdaEnergiaTrafoA4A4_kWh\tPerdaEnergiaFerroA4B_kWh\tPerdaEnergiaTrafoA4B_kWh\tPerdaEnergiaFerroA4A3a_kWh\tPerdaEnergiaTrafoA4A3a_kWh\tPerdaEnergiaFerroBA3a_kWh\tPerdaEnergiaTrafoBA3a_kWh\tPerdaEnergiaFerroBA4_kWh\tPerdaEnergiaTrafoBA4_kWh\tPerdaEnergiaLinhas_kWh\tPerdaEnergiaMT_kWh\tPerdaEnergiaBT_kWh\tPerdaEnergiaLinhasMT_kWh\tPerdaEnergiaLinhasBT_kWh\tPerdaEnergiaLinhasMTA3a_kWh\tPerdaEnergiaLinhasMTA4_kWh\tEnergiaForncMT_kWh\tEnergiaForncBT_kWh");
            mthdInputMensageLog("\t" + dblTotalPower_kW.ToString() + "\t" + dblTotalEnergy_kWh.ToString() + "\t" + dblTotalLoadEnergy_kWh.ToString() + "\t" + dblTecnicalLossPower_kW.ToString() + "\t" + dblTecnicalLossEnergy_kWh.ToString() + "\t" + dblNoLoadLosses_kWh.ToString() + "\t" + dblTransformerLosses_kWh.ToString() + "\t" + dblA3aA3aTransformerNoLoadLosses_kWh.ToString() + "\t" + dblA3aA3aTransformerLosses_kWh.ToString() + "\t" + dblA3aA4TransformerNoLoadLosses_kWh.ToString() + "\t" + dblA3aA4TransformerLosses_kWh.ToString() + "\t" + dblA3aBTransformerNoLoadLosses_kWh.ToString() + "\t" + dblA3aBTransformerLosses_kWh.ToString() + "\t" + dblA4A4TransformerNoLoadLosses_kWh.ToString() + "\t" + dblA4A4TransformerLosses_kWh.ToString() + "\t" + dblA4BTransformerNoLoadLosses_kWh.ToString() + "\t" + dblA4BTransformerLosses_kWh.ToString() + "\t" + dblA4A3aTransformerNoLoadLosses_kWh.ToString() + "\t" + dblA4A3aTransformerLosses_kWh.ToString() + "\t" + dblBA3aTransformerNoLoadLosses_kWh.ToString() + "\t" + dblBA3aTransformerLosses_kWh.ToString() + "\t" + dblBA4TransformerNoLoadLosses_kWh.ToString() + "\t" + dblBA4TransformerLosses_kWh.ToString() + "\t" + dblLineLosses_kWh.ToString() + "\t" + dblMTLosses_kWh.ToString() + "\t" + dblBTLosses_kWh.ToString() + "\t" + dblLineMTLosses_kWh.ToString() + "\t" + dblLineBTLosses_kWh.ToString() + "\t" + dblLineA3aLosses_kWh.ToString() + "\t" + dblLineA4Losses_kWh.ToString() + "\t" + dblMTLoad_kWh.ToString() + "\t" + dblBTLoad_kWh.ToString());
        }
        else
        {
            intConverged = 0;
            mthdInputMensageLog("Solução não convergiu.");
        }
    }

    private void mthdSolveCircuit()
    {
        if (_dss.Solution.Converged)
        {
            intConverged = 1;
            mthdInputMensageLog("Solução convergiu.");
        }
        else
        {
            intConverged = 0;
            mthdInputMensageLog("Solução não convergiu.");
        }
    }


    private void getSequenceVoltages()
    {
        mthdInputMensageLog(_dss.getSequenceVoltages());
    }

    private void getVoltages()
    {

        mthdInputMensageLog(_dss.getVoltages());
    }

    private void getMagnitudeAngle()
    {
        mthdInputMensageLog(_dss.getMagnitudeAngle());
    }

    private void getSequenceCurrents()
    {
        mthdInputMensageLog(_dss.getSequenceCurrents());
    }

    private void getLoadPowers()
    {
        mthdInputMensageLog(_dss.getLoadPowers());
    }

    private void mthdGetInfoMeters()
    {
        int i = 0;

        dblTotalPower_kW = 0;
        dblTotalEnergy_kWh = 0;
        dblTecnicalLossPower_kW = 0;
        dblTecnicalLossEnergy_kWh = 0;
        dblTotalLoadEnergy_kWh = 0;

        dblNoLoadLosses_kWh = 0;
        dblLineLosses_kWh = 0;

        dblTransformerLosses_kWh = 0;
        dblA3aA3aTransformerLosses_kWh = 0;
        dblA3aA4TransformerLosses_kWh = 0;
        dblA4A3aTransformerLosses_kWh = 0;
        dblA3aBTransformerLosses_kWh = 0;
        dblBA3aTransformerLosses_kWh = 0;
        dblA4A4TransformerLosses_kWh = 0;
        dblA4BTransformerLosses_kWh = 0;
        dblBA4TransformerLosses_kWh = 0;
        dblA3aA3aTransformerNoLoadLosses_kWh = 0;
        dblA3aA4TransformerNoLoadLosses_kWh = 0;
        dblA4A3aTransformerNoLoadLosses_kWh = 0;
        dblA3aBTransformerNoLoadLosses_kWh = 0;
        dblBA3aTransformerNoLoadLosses_kWh = 0;
        dblA4A4TransformerNoLoadLosses_kWh = 0;
        dblA4BTransformerNoLoadLosses_kWh = 0;
        dblBA4TransformerNoLoadLosses_kWh = 0;

        dblMTLosses_kWh = 0;
        dblA3aLosses_kWh = 0;
        dblA4Losses_kWh = 0;
        dblBTLosses_kWh = 0;

        dblLineMTLosses_kWh = 0;
        dblLineA3aLosses_kWh = 0;
        dblLineA4Losses_kWh = 0;
        dblLineBTLosses_kWh = 0;

        dblMTLoad_kWh = 0;
        dblBTLoad_kWh = 0;

        foreach (string strNomeMedidor in _dss.Circuit.Meters.AllNames)
        {
            _dss.Circuit.Meters.Name = strNomeMedidor;
            var dssMeter = _dss.Circuit.Meters;
            i = 0;
            foreach (string strRegister in dssMeter.RegisterNames)
            {
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 6) == "busa3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 6) == "busa4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta3aa3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA3aA3aTransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA3aA3aTransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta3aa4-")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA3aA4TransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA3aA4TransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta3ab--")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA3aBTransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA3aBTransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta4-a3a")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA4A3aTransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA4A3aTransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "tb--a3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblBA3aTransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblBA3aTransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta4-a4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA4A4TransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA4A4TransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "ta4-b--")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblA4BTransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblA4BTransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "tb--a4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "Transformer Losses")
                        dblBA4TransformerLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "No Load Losses kWh")
                        dblBA4TransformerNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "la3aa3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "la3aa4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "la4-a3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "lb--a3a")
                {
                    if (dssMeter.RegisterNames[i] == "44 kV Losses" || dssMeter.RegisterNames[i] == "40 kV Losses" || dssMeter.RegisterNames[i] == "34.5 kV Losses" || dssMeter.RegisterNames[i] == "33 kV Losses")
                        dblA3aLosses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "44 kV Line Loss" || dssMeter.RegisterNames[i] == "40 kV Line Loss" || dssMeter.RegisterNames[i] == "34.5 kV Line Loss" || dssMeter.RegisterNames[i] == "33 kV Line Loss")
                        dblLineA3aLosses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "la4-a4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                }
                if (GeoPerdasTools.mthdLeftString(dssMeter.Name, 7) == "lb--a4-")
                {
                    if (dssMeter.RegisterNames[i] == "25 kV Losses" || dssMeter.RegisterNames[i] == "24.2 kV Losses" || dssMeter.RegisterNames[i] == "23.1 kV Losses" || dssMeter.RegisterNames[i] == "23 kV Losses" || dssMeter.RegisterNames[i] == "22 kV Losses" || dssMeter.RegisterNames[i] == "13.8 kV Losses" || dssMeter.RegisterNames[i] == "13.2 kV Losses" || dssMeter.RegisterNames[i] == "11.95 kV Losses" || dssMeter.RegisterNames[i] == "11.9 kV Losses" || dssMeter.RegisterNames[i] == "11.8 kV Losses" || dssMeter.RegisterNames[i] == "11.4 kV Losses" || dssMeter.RegisterNames[i] == "11 kV Losses" || dssMeter.RegisterNames[i] == "6.6 kV Losses" || dssMeter.RegisterNames[i] == "6.3 kV Losses" || dssMeter.RegisterNames[i] == "3.8 kV Losses")
                        dblA4Losses_kWh += dssMeter.RegisterValues[i];
                    if (dssMeter.RegisterNames[i] == "25 kV Line Loss" || dssMeter.RegisterNames[i] == "24.2 kV Line Loss" || dssMeter.RegisterNames[i] == "23.1 kV Line Loss" || dssMeter.RegisterNames[i] == "23 kV Line Loss" || dssMeter.RegisterNames[i] == "22 kV Line Loss" || dssMeter.RegisterNames[i] == "13.8 kV Line Loss" || dssMeter.RegisterNames[i] == "13.2 kV Line Loss" || dssMeter.RegisterNames[i] == "11.95 kV Line Loss" || dssMeter.RegisterNames[i] == "11.9 kV Line Loss" || dssMeter.RegisterNames[i] == "11.8 kV Line Loss" || dssMeter.RegisterNames[i] == "11.4 kV Line Loss" || dssMeter.RegisterNames[i] == "11 kV Line Loss" || dssMeter.RegisterNames[i] == "6.6 kV Line Loss" || dssMeter.RegisterNames[i] == "6.3 kV Line Loss" || dssMeter.RegisterNames[i] == "3.8 kV Line Loss")
                        dblLineA4Losses_kWh += dssMeter.RegisterValues[i];
                }

                if (strRegister == "Zone Max kW")
                    dblTotalPower_kW += dssMeter.RegisterValues[i];
                if (strRegister == "Zone kWh" || strRegister == "Zone Losses kWh")
                    dblTotalEnergy_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "Zone Max kW Losses")
                    dblTecnicalLossPower_kW += dssMeter.RegisterValues[i];
                if (strRegister == "Zone Losses kWh")
                    dblTecnicalLossEnergy_kWh += dssMeter.RegisterValues[i];
                if (i == 60 || i == 61 || i == 62 || i == 63 || i == 64 || i == 65 || i == 66)
                    dblTotalLoadEnergy_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "No Load Losses kWh")
                    dblNoLoadLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "Line Losses")
                    dblLineLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "Transformer Losses")
                    dblTransformerLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "44 kV Losses" || strRegister == "40 kV Losses" || strRegister == "34.5 kV Losses" || strRegister == "33 kV Losses" || strRegister == "25 kV Losses" || strRegister == "24.2 kV Losses" || strRegister == "23.1 kV Losses" || strRegister == "23 kV Losses" || strRegister == "22 kV Losses" || strRegister == "13.8 kV Losses" || strRegister == "13.2 kV Losses" || strRegister == "11.95 kV Losses" || strRegister == "11.9 kV Losses" || strRegister == "11.8 kV Losses" || strRegister == "11.4 kV Losses" || strRegister == "11 kV Losses" || strRegister == "6.6 kV Losses" || strRegister == "6.3 kV Losses" || strRegister == "3.8 kV Losses")
                    dblMTLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "0.44 kV Losses" || strRegister == "0.4 kV Losses" || strRegister == "0.38 kV Losses" || strRegister == "0.254 kV Losses" || strRegister == "0.24 kV Losses" || strRegister == "0.231 kV Losses" || strRegister == "0.23 kV Losses" || strRegister == "0.22 kV Losses" || strRegister == "0.19 kV Losses" || strRegister == "0.127 kV Losses" || strRegister == "0.11 kV Losses")
                    dblBTLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "44 kV Line Loss" || strRegister == "40 kV Line Loss" || strRegister == "34.5 kV Line Loss" || strRegister == "33 kV Line Loss" || strRegister == "25 kV Line Loss" || strRegister == "24.2 kV Line Loss" || strRegister == "23.1 kV Line Loss" || strRegister == "23 kV Line Loss" || strRegister == "22 kV Line Loss" || strRegister == "13.8 kV Line Loss" || strRegister == "13.2 kV Line Loss" || strRegister == "11.95 kV Line Loss" || strRegister == "11.9 kV Line Loss" || strRegister == "11.8 kV Line Loss" || strRegister == "11.4 kV Line Loss" || strRegister == "11 kV Line Loss" || strRegister == "6.6 kV Line Loss" || strRegister == "6.3 kV Line Loss" || strRegister == "3.8 kV Line Loss")
                    dblLineMTLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "0.44 kV Line Loss" || strRegister == "0.4 kV Line Loss" || strRegister == "0.38 kV Line Loss" || strRegister == "0.254 kV Line Loss" || strRegister == "0.24 kV Line Loss" || strRegister == "0.231 kV Line Loss" || strRegister == "0.23 kV Line Loss" || strRegister == "0.22 kV Line Loss" || strRegister == "0.19 kV Line Loss" || strRegister == "0.127 kV Line Loss" || strRegister == "0.11 kV Line Loss")
                    dblLineBTLosses_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "44 kV Load Energy" || strRegister == "40 kV Load Energy" || strRegister == "34.5 kV Load Energy" || strRegister == "33 kV Load Energy" || strRegister == "25 kV Load Energy" || strRegister == "24.2 kV Load Energy" || strRegister == "23.1 kV Load Energy" || strRegister == "23 kV Load Energy" || strRegister == "22 kV Load Energy" || strRegister == "13.8 kV Load Energy" || strRegister == "13.2 kV Load Energy" || strRegister == "11.95 kV Load Energy" || strRegister == "11.9 kV Load Energy" || strRegister == "11.8 kV Load Energy" || strRegister == "11.4 kV Load Energy" || strRegister == "11 kV Load Energy" || strRegister == "6.6 kV Load Energy" || strRegister == "6.3 kV Load Energy" || strRegister == "3.8 kV Load Energy")
                    dblMTLoad_kWh += dssMeter.RegisterValues[i];
                if (strRegister == "0.44 kV Load Energy" || strRegister == "0.4 kV Load Energy" || strRegister == "0.38 kV Load Energy" || strRegister == "0.254 kV Load Energy" || strRegister == "0.24 kV Load Energy" || strRegister == "0.231 kV Load Energy" || strRegister == "0.23 kV Load Energy" || strRegister == "0.22 kV Load Energy" || strRegister == "0.19 kV Load Energy" || strRegister == "0.127 kV Load Energy" || strRegister == "0.11 kV Load Energy")
                    dblBTLoad_kWh += dssMeter.RegisterValues[i];
                i++;
            }
        }
    }

    private void mthdGetDSSCommandCircuit(string strName, double dblBase_kV, double dblOperation_pu, string strBus, int intConstraintOperation)
    {
        mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandCircuit(strName, dblBase_kV, dblOperation_pu, strBus, intConstraintOperation));
    }

    private string mthdWriteLanguageDSSCommandLinecode(string strName, double dblResis_ohms_km, double dblReat_ohms_km, double dblCorrMax_A)
    {
        string result;

        result = "New \"Linecode." + strName + "_1\" nphases=1 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", ".") + Environment.NewLine;
        result += "New \"Linecode." + strName + "_2\" nphases=2 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", ".") + Environment.NewLine;
        result += "New \"Linecode." + strName + "_3\" nphases=3 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", ".");

        return result;
    }

    private void mthdGetDSSCommandLinecode(string strName, double dblResis_ohms_km, double dblReat_ohms_km, double dblCorrMax_A)
    {
        mthdInsertCommandInLine("New \"Linecode." + strName + "_1\" nphases=1 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", "."));
        mthdInsertCommandInLine("New \"Linecode." + strName + "_2\" nphases=2 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", "."));
        mthdInsertCommandInLine("New \"Linecode." + strName + "_3\" nphases=3 basefreq=60 r1=" + dblResis_ohms_km.ToString().Replace(",", ".") + " x1=" + dblReat_ohms_km.ToString().Replace(",", ".") + " units=km normamps=" + dblCorrMax_A.ToString().Replace(",", "."));
    }

    private void mthdGetDSSCommandMTSwitch(string strName, string strCodFas, string strBus1, string strBus2)
    {
        mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandMTSwitch(strName, strCodFas, strBus1, strBus2));
    }

    private void mthdGetDSSCommandLine(string strName, string strCodFas, string strBus1, string strBus2, string strCodCond, double dblComp_km, string strProperty, int intConstraintLength, int intConstraintProperty)
    {
        mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandLine(strName, strCodFas, strBus1, strBus2, strCodCond, dblComp_km, strProperty, intConstraintLength, intConstraintProperty));
    }

    private string mthdWriteLanguageDSSCommandRegulatorMT(string strName, int intCodBnc, int intTipRegul, string strCodFasPrim, string strCodFasSecu, string strCodCoinc, string strBus1, string strBus2, double dblTensaoPrimTrafo_kV, double PotNom_kVA, double dblTenRgl_pu, double ReatHL_per, double dblPerdTtl_per, double dblPerdVz_per)
    {
        string result;

        result = "New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesRegulators(intTipRegul, strCodFasPrim) + " windings=2 buses=[\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodFasPrim) + "\" \"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodFasSecu) + "\"] conns=[" + mthdConexTrafo(strCodFasPrim) + " " + mthdConexTrafo(strCodFasSecu) + "] kvs=[" + mthdTensaoEnrolamento(strCodFasPrim, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + " " + mthdTensaoEnrolamento(strCodFasSecu, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + "] kvas=[" + PotNom_kVA.ToString().Replace(",", ".") + " " + PotNom_kVA.ToString().Replace(",", ".") + "] xhl=" + ReatHL_per.ToString().Replace(",", ".") + " %loadloss=" + (dblPerdTtl_per - dblPerdVz_per).ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVz_per.ToString().Replace(",", ".") + Environment.NewLine;
        result += "New \"Regcontrol." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" transformer=\"" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" winding=2 vreg=" + (dblTenRgl_pu * 100).ToString().Replace(",", ".") + " band=2 ptratio=" + (mthdTensaoEnrolamento(strCodFasSecu, dblTensaoPrimTrafo_kV) * 10).ToString().Replace(",", ".");
        if (intTipRegul == 2 && intCodBnc == 2)
            result += Environment.NewLine + "New \"Line." + strName + "_J\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodCoinc) + " bus1=\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodCoinc) + "\" bus2=\"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodCoinc) + "\" r0=1e-3  r1=1e-3  x0=0  x1=0  c0=0 c1=0";

        return result;
    }

    private void mthdGetDSSCommandRegulatorMT(string strName, int intCodBnc, int intTipRegul, string strCodFasPrim, string strCodFasSecu, string strCodCoinc, string strBus1, string strBus2, double dblTensaoPrimTrafo_kV, double PotNom_kVA, double dblTenRgl_pu, double ReatHL_per, double dblPerdTtl_per, double dblPerdVz_per)
    {
        mthdInsertCommandInLine("New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesRegulators(intTipRegul, strCodFasPrim) + " windings=2 buses=[\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodFasPrim) + "\" \"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodFasSecu) + "\"] conns=[" + mthdConexTrafo(strCodFasPrim) + " " + mthdConexTrafo(strCodFasSecu) + "] kvs=[" + mthdTensaoEnrolamento(strCodFasPrim, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + " " + mthdTensaoEnrolamento(strCodFasSecu, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + "] kvas=[" + PotNom_kVA.ToString().Replace(",", ".") + " " + PotNom_kVA.ToString().Replace(",", ".") + "] xhl=" + ReatHL_per.ToString().Replace(",", ".") + " %loadloss=" + (dblPerdTtl_per - dblPerdVz_per).ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVz_per.ToString().Replace(",", "."));
        mthdInsertCommandInLine("New \"Regcontrol." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" transformer=\"" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" winding=2 vreg=" + (dblTenRgl_pu * 100).ToString().Replace(",", ".") + " band=2 ptratio=" + (mthdTensaoEnrolamento(strCodFasSecu, dblTensaoPrimTrafo_kV) * 10).ToString().Replace(",", "."));
        if (intTipRegul == 2 && intCodBnc == 2)
            mthdInsertCommandInLine("New \"Line." + strName + "_J\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodCoinc) + " bus1=\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodCoinc) + "\" bus2=\"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodCoinc) + "\" r0=1e-3  r1=1e-3  x0=0  x1=0  c0=0 c1=0");
    }

    private string mthdConexTrafo(string strCodFas)
    {
        if (strCodFas == "A" || strCodFas == "B" || strCodFas == "C" || strCodFas == "AN" || strCodFas == "BN" || strCodFas == "CN" || strCodFas == "ABN" || strCodFas == "BCN" || strCodFas == "CAN" || strCodFas == "ABCN")
            return "Wye";
        else if (strCodFas == "AB" || strCodFas == "BC" || strCodFas == "CA" || strCodFas == "ABC")
            return "Delta";
        return "";
    }

    private double mthdTensaoEnrolamento(string strCodFas, double dblTensao_kV)
    {
        if (strCodFas == "A" || strCodFas == "B" || strCodFas == "C" || strCodFas == "AN" || strCodFas == "BN" || strCodFas == "CN" || strCodFas == "ABN" || strCodFas == "BCN" || strCodFas == "CAN" || strCodFas == "ABCN")
            return dblTensao_kV / Math.Pow(3, 0.5);
        else if (strCodFas == "AB" || strCodFas == "BC" || strCodFas == "CA" || strCodFas == "ABC")
            return dblTensao_kV;
        return 0.0;
    }

    private double mthdTensaoEnrolamentoSecundario(string strCodFas, double dblTensao_kV)
    {
        if (strCodFas == "A" || strCodFas == "B" || strCodFas == "C" || strCodFas == "AN" || strCodFas == "BN" || strCodFas == "CN")
            return dblTensao_kV;
        else if (strCodFas == "ABN" || strCodFas == "BCN" || strCodFas == "CAN" || strCodFas == "ABCN")
            return dblTensao_kV / Math.Pow(3, 0.5);
        else if (strCodFas == "AB" || strCodFas == "BC" || strCodFas == "CA" || strCodFas == "ABC")
            return dblTensao_kV;
        return 0.0;
    }

    private void mthdGetDSSCommandLoadshape(string strName, double PotAtvNorm01, double PotAtvNorm02, double PotAtvNorm03, double PotAtvNorm04, double PotAtvNorm05, double PotAtvNorm06, double PotAtvNorm07, double PotAtvNorm08, double PotAtvNorm09, double PotAtvNorm10, double PotAtvNorm11, double PotAtvNorm12, double PotAtvNorm13, double PotAtvNorm14, double PotAtvNorm15, double PotAtvNorm16, double PotAtvNorm17, double PotAtvNorm18, double PotAtvNorm19, double PotAtvNorm20, double PotAtvNorm21, double PotAtvNorm22, double PotAtvNorm23, double PotAtvNorm24)
    {
        mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandLoadshape(strName, PotAtvNorm01, PotAtvNorm02, PotAtvNorm03, PotAtvNorm04, PotAtvNorm05, PotAtvNorm06, PotAtvNorm07, PotAtvNorm08, PotAtvNorm09, PotAtvNorm10, PotAtvNorm11, PotAtvNorm12, PotAtvNorm13, PotAtvNorm14, PotAtvNorm15, PotAtvNorm16, PotAtvNorm17, PotAtvNorm18, PotAtvNorm19, PotAtvNorm20, PotAtvNorm21, PotAtvNorm22, PotAtvNorm23, PotAtvNorm24));
    }

    private string mthdWriteLanguageDSSCommandTransformer(string strName, int intCodBnc, int intMRT, int intTipTrafo, string strCodFasPrim, string strCodFasSecu, string strCodFasTerc, string strBus1, string strBus2, double dblTensaoPrimTrafo_kV, double dblTensaoSecuTrafo_kV, double dblTap_pu, double dblPotNom_kVA, double dblPerdTtl_per, double dblPerdVz_per, string strPropriedade, int intSemCarga)
    {
        string result = "";
        double dblPerdTtlAlt_per = dblPerdTtl_per, dblPerdVzAlt_per = dblPerdVz_per;
        double dblPerdLodAlt_per = dblPerdTtl_per - dblPerdVz_per;

        if (strPropriedade == "TC" && intNeutralizarTrafoTerceiros == 1)
        {
            dblPerdTtlAlt_per = 0;
            dblPerdVzAlt_per = 0;
            dblPerdLodAlt_per = 0;
        }
        if (intSemCarga == 0)
            if (intMRT == 1)
            {
                result += "New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", ".");
                result += Environment.NewLine + mthdWriteLanguageDSSCommandLinecode("LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 15E3, 0, 0);
                result += Environment.NewLine + GeoPerdasTools.mthdWriteLanguageDSSCommandLine("Resist_MTR_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strCodFasPrim, strBus1, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), "LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 1E-3, "PR", 0, 0);
            }
            else
                result += "New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, strBus1, strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", ".");
        else
            if (intMRT == 1)
        {
            result += "!New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", ".");
            result += Environment.NewLine + "!" + mthdWriteLanguageDSSCommandLinecode("LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 15E3, 0, 0);
            result += Environment.NewLine + "!" + GeoPerdasTools.mthdWriteLanguageDSSCommandLine("Resist_MTR_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strCodFasPrim, strBus1, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), "ABC", 1E-3, "PR", 0, 0);
        }
        else
            result += "!New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, strBus1, strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", ".");

        return result;
    }

    private void mthdGetDSSCommandTransformer(string strName, int intCodBnc, int intMRT, int intTipTrafo, string strCodFasPrim, string strCodFasSecu, string strCodFasTerc, string strBus1, string strBus2, double dblTensaoPrimTrafo_kV, double dblTensaoSecuTrafo_kV, double dblTap_pu, double dblPotNom_kVA, double dblPerdTtl_per, double dblPerdVz_per, string strPropriedade, int intSemCarga)
    {
        double dblPerdTtlAlt_per = dblPerdTtl_per, dblPerdVzAlt_per = dblPerdVz_per;
        double dblPerdLodAlt_per = dblPerdTtl_per - dblPerdVz_per;

        if (strPropriedade == "TC" && intNeutralizarTrafoTerceiros == 1)
        {
            dblPerdTtlAlt_per = 0;
            dblPerdVzAlt_per = 0;
            dblPerdLodAlt_per = 0;
        }
        if (intSemCarga == 0)
            if (intMRT == 1)
            {
                mthdInsertCommandInLine("New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", "."));
                mthdGetDSSCommandLinecode("LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 15E3, 0, 0);
                mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandLine("Resist_MTR_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strCodFasPrim, strBus1, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), "LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 1E-3, "PR", 0, 0));
            }
            else
                mthdInsertCommandInLine("New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, strBus1, strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", "."));
        else
            if (intMRT == 1)
        {
            mthdInsertCommandInLine("!New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", "."));
            mthdGetDSSCommandLinecode("LC_MRT_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), 15E3, 0, 0);
            mthdInsertCommandInLine("!" + GeoPerdasTools.mthdWriteLanguageDSSCommandLine("Resist_MTR_" + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), strCodFasPrim, strBus1, "MRT_" + strBus1 + strName + GeoPerdasTools.mthdNamingBank(intCodBnc), "ABC", 1E-3, "PR", 0, 0));
        }
        else
            mthdInsertCommandInLine("!New \"Transformer." + strName + GeoPerdasTools.mthdNamingBank(intCodBnc) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesTransformers(strCodFasPrim) + " windings=" + mthdEnrrolamentoTrafo(strCodFasTerc, strCodFasSecu) + " buses=[" + mthdBusesTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc, strBus1, strBus2) + "] conns=[" + mthdConnsTrafo(strCodFasPrim, strCodFasSecu, strCodFasTerc) + "] kvs=[" + mthdkVsTrafo(intTipTrafo, strCodFasPrim, strCodFasSecu, strCodFasTerc, dblTensaoPrimTrafo_kV, dblTensaoSecuTrafo_kV) + "] " + (intAdequarTapTrafo == 1 ? "taps=[" + mthdTapsTrafo(strCodFasTerc, dblTap_pu) + "] " : "") + "kvas=[" + mthdkVAsTrafo(strCodFasSecu, strCodFasTerc, dblPotNom_kVA) + "] %loadloss=" + dblPerdLodAlt_per.ToString().Replace(",", ".") + " %noloadloss=" + dblPerdVzAlt_per.ToString().Replace(",", "."));
    }

    private string mthdEnrrolamentoTrafo(string strCodFas1, string strCodFas2)
    {
        if (strCodFas1 == "XX" && strCodFas2 != "ABN")
            return "2";
        else if (strCodFas1 == "BN" || strCodFas1 == "CN" || strCodFas1 == "AN" || strCodFas2 == "ABN")
            return "3";
        return "0";
    }

    private string mthdBusesTrafo(string strCodFas1, string strCodFas2, string strCodFas3, string strBus1, string strBus2)
    {
        if (strCodFas3 == "BN" || strCodFas3 == "CN" || strCodFas3 == "AN" || strCodFas2 == "ABN")
            return "\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodFas1) + "\" \"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodFas2) + "\" \"" + strBus2 + GeoPerdasTools.mthdPhasingTertiary(strCodFas3) + "\"";
        else if (strCodFas3 == "XX" || strCodFas2 == "AN" || strCodFas2 == "BN" || strCodFas2 == "CN" || strCodFas2 == "AB" || strCodFas2 == "BC" || strCodFas2 == "CA" || strCodFas2 == "AC" || strCodFas2 == "ABCN" || strCodFas2 == "ABC")
            return "\"" + strBus1 + GeoPerdasTools.mthdPhasing(strCodFas1) + "\" \"" + strBus2 + GeoPerdasTools.mthdPhasing(strCodFas2) + "\"";
        return "";
    }

    private string mthdConnsTrafo(string strCodFas1, string strCodFas2, string strCodFas3)
    {
        if (strCodFas3 == "BN" || strCodFas3 == "CN" || strCodFas3 == "AN" || strCodFas2 == "ABN")
            return mthdConexTrafo(strCodFas1) + " " + mthdConexTrafo(strCodFas2) + " " + mthdConexTrafo(strCodFas3);
        else if (strCodFas3 == "XX" || strCodFas2 == "AN" || strCodFas2 == "BN" || strCodFas2 == "CN" || strCodFas2 == "AB" || strCodFas2 == "BC" || strCodFas2 == "CA" || strCodFas2 == "AC" || strCodFas2 == "ABCN" || strCodFas2 == "ABC")
            return mthdConexTrafo(strCodFas1) + " " + mthdConexTrafo(strCodFas2);
        return "";
    }

    private string mthdkVsTrafo(int intTipTrafo, string strCodFas1, string strCodFas2, string strCodFas3, double dblTensaoPrimTrafo_kV, double dblTensaoSecuTrafo_kV)
    {
        if (intTipTrafo == 4)
            return dblTensaoPrimTrafo_kV.ToString().Replace(",", ".") + " " + dblTensaoSecuTrafo_kV.ToString().Replace(",", ".");
        else if (strCodFas3 == "BN" || strCodFas3 == "CN" || strCodFas3 == "AN" || strCodFas2 == "ABN")
            return mthdTensaoEnrolamento(strCodFas1, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + " " + (dblTensaoSecuTrafo_kV / 2).ToString().Replace(",", ".") + " " + (dblTensaoSecuTrafo_kV / 2).ToString().Replace(",", ".");
        else if (strCodFas3 == "XX" && (strCodFas2 == "AN" || strCodFas2 == "BN" || strCodFas2 == "CN" || strCodFas2 == "AB" || strCodFas2 == "BC" || strCodFas2 == "CA" || strCodFas2 == "AC" || strCodFas2 == "ABC"))
            return mthdTensaoEnrolamento(strCodFas1, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + " " + mthdTensaoEnrolamentoSecundario(strCodFas2, dblTensaoSecuTrafo_kV).ToString().Replace(",", ".");
        else if (strCodFas3 == "XX" && strCodFas2 == "ABCN")
            return mthdTensaoEnrolamento(strCodFas1, dblTensaoPrimTrafo_kV).ToString().Replace(",", ".") + " " + dblTensaoSecuTrafo_kV.ToString().Replace(",", ".");
        return "";
    }

    private string mthdkVAsTrafo(string strCodFas2, string strCodFas3, double dblPotNom_kVA)
    {
        if (strCodFas3 == "BN" || strCodFas3 == "CN" || strCodFas3 == "AN" || strCodFas2 == "ABN")
            return dblPotNom_kVA.ToString().Replace(",", ".") + " " + dblPotNom_kVA.ToString().Replace(",", ".") + " " + dblPotNom_kVA.ToString().Replace(",", ".");
        else if (strCodFas3 == "XX" || strCodFas2 == "AN" || strCodFas2 == "BN" || strCodFas2 == "CN" || strCodFas2 == "AB" || strCodFas2 == "BC" || strCodFas2 == "CA" || strCodFas2 == "AC" || strCodFas2 == "ABCN" || strCodFas2 == "ABC")
            return dblPotNom_kVA.ToString().Replace(",", ".") + " " + dblPotNom_kVA.ToString().Replace(",", ".");
        return "";
    }

    private string mthdTapsTrafo(string strCodFas3, double dblTap_pu)
    {
        if (strCodFas3 != "XX")
            return "1 " + dblTap_pu.ToString().Replace(",", ".") + " " + dblTap_pu.ToString().Replace(",", ".");
        else
            return "1 " + dblTap_pu.ToString().Replace(",", ".");
    }

    private string mthdWriteLanguageDSSCommandNewLoad(string strName, string strCodFas, string strBus, double dblTensaoLine_kV, double dblTensaoPhase_kV, double dblDemMax_kW, double dblDemMaxTrafo_kW, string strCodCrvCrg, int intTipoCarga, int intAlteraModelo, int intExecucaoSnapShot)
    {
        string result = "";
        double dblDemMaxCorrigida_kW;
        string strCurvaCarga;

        if (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW)
            dblDemMaxCorrigida_kW = dblDemMaxTrafo_kW;
        else
            dblDemMaxCorrigida_kW = dblDemMax_kW;

        if (intExecucaoSnapShot == 1)
            strCurvaCarga = "";
        else
            strCurvaCarga = " daily=\"" + strCodCrvCrg + "\" status=variable";

        if (intTipoCarga == 1)
        {
            if (intAlteraModelo == 1)
            {
                result = "New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.93" + Environment.NewLine;
                result += "New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.93";
            }
            else
            {
                result = "New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", ".") + Environment.NewLine;
                result += "New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", ".");
            }
        }
        else if (intTipoCarga == 2)
        {
            if (intAlteraModelo == 1)
            {
                result = "New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.92" + (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW && intAdequarPotenciaCarga == 1 ? " ! Carga limitada" : "") + Environment.NewLine;
                result += "New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.92" + (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW && intAdequarPotenciaCarga == 1 ? " ! Carga limitada" : "");
            }
            else
            {
                result = "New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", ".") + (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW && intAdequarPotenciaCarga == 1 ? " ! Carga limitada" : "") + Environment.NewLine;
                result += "New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", ".") + (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW && intAdequarPotenciaCarga == 1 ? " ! Carga limitada" : "");
            }
        }

        return result;
    }

    private string mthdGetTypeModel(int intPart)
    {
        if (intAdequarModeloCarga == 1)
            return intPart == 1 ? "2" : "3";
        else if (intAdequarModeloCarga == 2)
            return intPart == 1 ? "1" : "1";
        else if (intAdequarModeloCarga == 3)
            return intPart == 1 ? "3" : "3";
        return "";
    }

    private void mthdGetDSSCommandNewLoad(string strName, string strCodFas, string strBus, double dblTensaoLine_kV, double dblTensaoPhase_kV, double dblDemMax_kW, double dblDemMaxTrafo_kW, string strCodCrvCrg, int intTipoCarga, int intAlteraModelo, int intExecucaoSnapShot)
    {
        double dblDemMaxCorrigida_kW;
        string strCurvaCarga;

        if (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW)
            dblDemMaxCorrigida_kW = dblDemMaxTrafo_kW;
        else
            dblDemMaxCorrigida_kW = dblDemMax_kW;

        if (intExecucaoSnapShot == 1)
            strCurvaCarga = "";
        else
            strCurvaCarga = " daily=\"" + strCodCrvCrg + "\" status=variable";

        if (intTipoCarga == 1)
        {
            if (intAlteraModelo == 1)
            {
                mthdInsertCommandInLine("New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.93");
                mthdInsertCommandInLine("New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.93");
            }
            else
            {
                mthdInsertCommandInLine("New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", "."));
                mthdInsertCommandInLine("New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", "."));
            }
        }
        else if (intTipoCarga == 2)
        {
            if (intAlteraModelo == 1)
            {
                mthdInsertCommandInLine("New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.92");
                mthdInsertCommandInLine("New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=0.92");
            }
            else
            {
                mthdInsertCommandInLine("New \"Load." + strName + "_M1\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(1) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", "."));
                mthdInsertCommandInLine("New \"Load." + strName + "_M2\" bus1=\"" + strBus + GeoPerdasTools.mthdPhasing(strCodFas) + "\" phases=" + GeoPerdasTools.mthdNumberPhasesLoads(strCodFas) + " conn=" + mthdTipoConexaoCargas(strCodFas) + " model=" + mthdGetTypeModel(2) + " kv=" + GeoPerdasTools.mthdTensionLoads(strCodFas, dblTensaoLine_kV, dblTensaoPhase_kV).ToString().Replace(",", ".") + " kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + " pf=0.92" + strCurvaCarga + " vmaxpu=1.5 vminpu=" + dblVPUMin.ToString().Replace(",", "."));
            }
        }
    }

    private string mthdTipoConexaoCargas(string strCodFas)
    {
        if (strCodFas == "A" || strCodFas == "B" || strCodFas == "C" || strCodFas == "AN" || strCodFas == "BN" || strCodFas == "CN")
            return "Wye";
        else if (strCodFas == "AB" || strCodFas == "BC" || strCodFas == "CA" || strCodFas == "ABN" || strCodFas == "BCN" || strCodFas == "CAN" || strCodFas == "ABC" || strCodFas == "ABCN")
            return "Delta";
        return "";
    }

    private void mthdGetDSSCommandEditLoad(string strName, string strCodFas, double dblDemMax_kW, double dblDemMaxTrafo_kW, string strCodCrvCrg, int intTipoCarga, int intExecucaoSnapShot)
    {
        double dblDemMaxCorrigida_kW;
        string strCurvaCarga;

        if (intTipoCarga == 2 && dblDemMax_kW > dblDemMaxTrafo_kW)
            dblDemMaxCorrigida_kW = dblDemMaxTrafo_kW;
        else
            dblDemMaxCorrigida_kW = dblDemMax_kW;

        if (intExecucaoSnapShot == 1)
            strCurvaCarga = "";
        else
            strCurvaCarga = " daily=\"" + strCodCrvCrg + "\"";

        mthdInsertCommandInLine("Edit \"Load." + strName + "_M1\" kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + strCurvaCarga);
        mthdInsertCommandInLine("Edit \"Load." + strName + "_M2\" kw=" + (intAdequarPotenciaCarga == 1 ? (dblDemMaxCorrigida_kW / 2).ToString().Replace(",", ".") : (dblDemMax_kW / 2).ToString().Replace(",", ".")) + strCurvaCarga);
    }

    private void mthdGetDSSCommandEnergymeter(string strName, string strElement)
    {
        mthdInsertCommandInLine(GeoPerdasTools.mthdWriteLanguageDSSCommandEnergymeter(strName, strElement));
    }

    private string mthdNomeSegmentoTensaoBus(double dblTensao_kV)
    {
        if (dblTensao_kV > 25)
            return "A3a";
        if (dblTensao_kV <= 25 && dblTensao_kV > 1)
            return "A4-";
        if (dblTensao_kV <= 1)
            return "B--";
        return "";
    }

    private string mthdNomeElemento(string strTipoElemento)
    {
        if (strTipoElemento == "RML")
            return "Line.RBT_";
        else if (strTipoElemento == "SEGMBT")
            return "Line.SBT_";
        else if (strTipoElemento == "SEGMMT")
            return "Line.SMT_";
        else if (strTipoElemento == "CHVMT")
            return "Line.CMT_";
        else if (strTipoElemento == "TRAFO")
            return "Transformer.TRF_";
        else if (strTipoElemento == "REGUL")
            return "Transformer.REG_";
        return "";
    }

    private void mthdWriteLineInFile(string Line)
    {
        //this.swFile.WriteLine(Line);
    }

    private void clbOptionsRun_SelectedIndexChanged(object sender, EventArgs e)
    {
        mthdUpdateOptionsRun();
    }

    private void mthdUpdateOptionsRun()
    {
        //intRealizaCnvrgcPNT = Convert.ToInt32(config.clbOption["Convergência de Perda Não Técnica"]);
        //intUsaTrafoABNT = Convert.ToInt32(config.clbOption["Transformadores ABNT"]);
        //intAdequarTensaoCargasMT = Convert.ToInt32(config.clbOption["Adequação Tensão Mínima das Cargas MT (0,93 pu)"]);
        //intAdequarTensaoCargasBT = Convert.ToInt32(config.clbOption["Adequação Tensão Mínima das Cargas BT (0,92 pu)"]);
        //intAdequarTensaoSuperior = Convert.ToInt32(config.clbOption["Limitar Máxima Tensão de Barras e Reguladores (1,05 pu)"]);
        //intAdequarRamal = Convert.ToInt32(config.clbOption["Limitar o Ramal (30m)"]);
        //intAdequarTapTrafo = Convert.ToInt32(config.clbOption["Utilizar Tap nos Transformadores"]);
        //intAdequarPotenciaCarga = Convert.ToInt32(config.clbOption["Limitar Cargas BT (Potência ativa do transformador)"]);
        //intAdequarTrafoVazio = Convert.ToInt32(config.clbOption["Eliminar Transformadores Vazios"]);
        //intNeutralizarTrafoTerceiros = Convert.ToInt32(config.clbOption["Neutralizar Transformadores de Terceiros"]);
        //intNeutralizarRedeTerceiros = Convert.ToInt32(config.clbOption["Neutralizar Redes de Terceiros (MT/BT)"]);

        intRealizaCnvrgcPNT = Convert.ToInt32(config.intRealizaCnvrgcPNT);
        intUsaTrafoABNT = Convert.ToInt32(config.transformadoresABNT);
        intAdequarTensaoCargasMT = Convert.ToInt32(config.adequaBT);
        intAdequarTensaoCargasBT = Convert.ToInt32(config.adequaMT);
        intAdequarTensaoSuperior = Convert.ToInt32(config.intAdequarTensaoSuperior);
        intAdequarRamal = Convert.ToInt32(config.intAdequarRamal);
        intAdequarTapTrafo = Convert.ToInt32(config.intAdequarTapTrafo);
        intAdequarPotenciaCarga = Convert.ToInt32(config.intAdequarPotenciaCarga);
        intAdequarTrafoVazio = Convert.ToInt32(config.trafoVazio);
        intNeutralizarTrafoTerceiros = Convert.ToInt32(config.neutralizaTrafo);
        intNeutralizarRedeTerceiros = Convert.ToInt32(config.netralizaRede);
        intAdequarModeloCarga = Convert.ToInt32(config.cbModel);
        if (config.tbVPUMin == "")
            config.tbVPUMin = "0,5";
        dblVPUMin = Convert.ToDouble(config.tbVPUMin.Replace(".", ","));

    }

    private string mthdGetOptionsToString()
    {
        string result = "";

        result += intRealizaCnvrgcPNT == 1 ? "N" : "-";
        result += intUsaTrafoABNT == 1 ? "T" : "-";
        result += intAdequarTensaoCargasMT == 1 ? "M" : "-";
        result += intAdequarTensaoCargasBT == 1 ? "B" : "-";
        result += intAdequarTensaoSuperior == 1 ? "S" : "-";
        result += intAdequarRamal == 1 ? "R" : "-";
        result += intAdequarModeloCarga == 1 ? "1" : intAdequarModeloCarga == 2 ? "2" : intAdequarModeloCarga == 3 ? "3" : "-";
        result += intAdequarPotenciaCarga == 1 ? "P" : "-";
        result += intAdequarTrafoVazio == 1 ? "V" : "-";
        result += intAdequarTapTrafo == 1 ? "T" : "-";
        result += intNeutralizarTrafoTerceiros == 1 ? "T" : "-";
        result += intNeutralizarRedeTerceiros == 1 ? "R" : "-";

        return result;
    }

    private void addToZipFile(IDictionary<string, string[]> dictLines, string nomeArquivo)
    {
        var lines = new List<string>();
        dictLines.TryAdd(nomeArquivo, this.strFileData.Where(s => !string.IsNullOrEmpty(s) && s != "<FIM DOS DADOS>").ToArray());       
    }

    public IDictionary<string, string[]> bExportDB_Click()
    {

        int i;
        String strPath;
        String strTensoesBase;

        var dictLines = new Dictionary<string, string[]>();

        if (this.sqlConnServer == null)
        {
            this.mthdInputMensageLog("A conexão com o banco de dados não foi estabelecida.");
        }
        else if (this.config.tbObjectBase == null || this.config.tbObjectBase == "")
        {
            this.mthdInputMensageLog("A base a ser calculada não foi indicada.");
        }
        else
        {
            try
            {

                //// Obtem a pasta onde os arquivos serão gravados
                //if (this.fbdSelectPathWrite.ShowDialog() == DialogResult.Cancel)
                //    return;
                //else
                //    strPath = this.fbdSelectPathWrite.SelectedPath;

                // Captura a base a ser calculada
                this.strBaseSelected = this.config.tbObjectBase;

                // Loop da base e dos alimentadores
                this.strFeedersSelected.Clear();
                if (this.config.tbObjectFeeder == null || this.config.tbObjectFeeder == "")
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + this.strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected, this.sqlConnServer);
                else
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + this.strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND " + GeoPerdasTools.mthdQuerySQLFeeders(this.config.tbObjectFeeder), this.sqlConnServer);
                this.sqlCommServer.CommandTimeout = 0;
                this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                while (this.sqlDtRdrServer.Read())
                {
                    this.strFeedersSelected.Add(this.sqlDtRdrServer["CodAlim"].ToString());
                }
                this.sqlDtRdrServer.Close();

                foreach (String strFeeder in strFeedersSelected)
                {

                    // Início do processo
                    this.mthdInputMensageLog("Início do processo para o alimentador " + strFeeder + ".");

                    // Verifica as opções solicitadas pelo usuário
                    this.mthdUpdateOptionsRun();

                    // Loop dos códigos dos circuitos
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + this.strBaseSelected + "CircMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de circuitos";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandCircuit(strFeeder, Convert.ToDouble(this.sqlDtRdrServer["TenNom_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TenOpe_pu"]), this.sqlDtRdrServer["CodPonAcopl"].ToString(), this.intAdequarTensaoSuperior);
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";
                    
                    this.addToZipFile(dictLines, @"\CircuitoMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelo circuito.");

                    // Loop dos códigos dos condutores
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + this.strBaseSelected + "CodCondutor] AS t1 WHERE t1.CodBase = " + this.strBaseSelected, this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de condutores";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = this.mthdWriteLanguageDSSCommandLinecode(this.sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["Resis_ohms_km"]), Convert.ToDouble(this.sqlDtRdrServer["Reat_ohms_km"]), Convert.ToDouble(this.sqlDtRdrServer["CorrMax_A"]));
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\CodCondutor_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelos condutores.");

                    // Loop dos códigos das curvas MT
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxCrvCrgMTHor AS t1 WHERE t1.CodBase = " + this.strBaseSelected, this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de curvas MT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandLoadshape(this.sqlDtRdrServer["CodCrvCrg"].ToString() + "_" + this.sqlDtRdrServer["TipoDia"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm01"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm02"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm03"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm04"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm05"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm06"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm07"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm08"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm09"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm10"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm11"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm12"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm13"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm14"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm15"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm16"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm17"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm18"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm19"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm20"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm21"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm22"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm23"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm24"]));
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";
                    
                    this.addToZipFile(dictLines, @"\CurvacargaMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelas curvas MT.");

                    // Loop dos códigos das curvas BT
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxCrvCrgBTHor AS t1 WHERE t1.CodBase = " + this.strBaseSelected, this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de curvas BT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandLoadshape(this.sqlDtRdrServer["CodCrvCrg"].ToString() + "_" + this.sqlDtRdrServer["TipoDia"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm01"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm02"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm03"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm04"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm05"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm06"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm07"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm08"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm09"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm10"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm11"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm12"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm13"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm14"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm15"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm16"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm17"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm18"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm19"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm20"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm21"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm22"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm23"]), Convert.ToDouble(this.sqlDtRdrServer["PotAtvNorm24"]));
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\CurvacargaBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");

                    this.mthdInputMensageLog("Passou pelas curvas BT.");

                    // Loop dos chaves MT
                    this.sqlCommServer = new SqlCommand("SELECT t1.CodChvMT, t1.CodFas, t1.De, t1.Para FROM dbo.[" + this.strBaseSelected + "ChaveMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "' AND t1.EstChv = 2", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de chaves MT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandMTSwitch("CMT_" + this.sqlDtRdrServer["CodChvMT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["De"].ToString(), this.sqlDtRdrServer["Para"].ToString());
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\ChavesMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                   
                    this.mthdInputMensageLog("Passou pelas chaves MT.");

                    // Loop dos segmentos MT
                    this.sqlCommServer = new SqlCommand("SELECT t1.CodSegmMT, t1.CodFas, t1.CodCond, t1.Comp_km, t1.De, t1.Para, t1.Propr FROM dbo.[" + this.strBaseSelected + "SegmentoMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de segmentos MT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandLine("SMT_" + this.sqlDtRdrServer["CodSegmMT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["De"].ToString(), this.sqlDtRdrServer["Para"].ToString(), this.sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["Comp_km"]), this.sqlDtRdrServer["Propr"].ToString(), 0, intNeutralizarRedeTerceiros);
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";
                    
                    this.addToZipFile(dictLines, @"\SegmentosMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelos segmentos MT.");

                    // Loop dos reguladores MT
                    this.sqlCommServer = new SqlCommand("SELECT t1.CodRegulMT, t1.CodBnc, t1.TipRegul, t1.CodFasPrim, t1.CodFasSecu, t1.PotNom_kVA, t1.TenRgl_pu, t1.[ReatHL_%], t1.PerdTtl_W, t1.PerdVz_W, t1.TnsLnh1_kV, t1.De, t1.Para, t1.CodFasCoinc FROM dbo.[" + this.strBaseSelected + "ReguladorMT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de reguladores MT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = this.mthdWriteLanguageDSSCommandRegulatorMT("REG_" + this.sqlDtRdrServer["CodRegulMT"].ToString(), Convert.ToInt32(this.sqlDtRdrServer["CodBnc"]), Convert.ToInt32(this.sqlDtRdrServer["TipRegul"]), this.sqlDtRdrServer["CodFasPrim"].ToString(), this.sqlDtRdrServer["CodFasSecu"].ToString(), this.sqlDtRdrServer["CodFasCoinc"].ToString(), this.sqlDtRdrServer["De"].ToString(), this.sqlDtRdrServer["Para"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"]), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(this.sqlDtRdrServer["TenRgl_pu"]), Convert.ToDouble(this.sqlDtRdrServer["ReatHL_%"]), Convert.ToDouble(this.sqlDtRdrServer["PerdTtl_W"]) / (1000 * Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"])), Convert.ToDouble(this.sqlDtRdrServer["PerdVz_W"]) / (1000 * Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"])));
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\ReguladorMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelos reguladores MT.");

                    // Loop dos transformadores MT-MT ou MT-BT
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.[" + this.strBaseSelected + "TrafoMTMTMTBT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de transformadores MT-MT ou MT-BT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = this.mthdWriteLanguageDSSCommandTransformer("TRF_" + this.sqlDtRdrServer["CodTrafo"].ToString(), Convert.ToInt32(this.sqlDtRdrServer["CodBnc"]), Convert.ToInt32(this.sqlDtRdrServer["MRT"]), Convert.ToInt32(this.sqlDtRdrServer["TipTrafo"]), this.sqlDtRdrServer["CodFasPrim"].ToString(), this.sqlDtRdrServer["CodFasSecu"].ToString(), this.sqlDtRdrServer["CodFasTerc"].ToString(), this.sqlDtRdrServer["De"].ToString(), this.sqlDtRdrServer["Para"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TenSecu_kV"]), Convert.ToDouble(this.sqlDtRdrServer["Tap_pu"]), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), (GeoPerdasTools.mthdTransformerTotalPowerLoss_per(this.sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"])) != 0 && this.intUsaTrafoABNT == 1) ? GeoPerdasTools.mthdTransformerTotalPowerLoss_per(this.sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"])) : Convert.ToDouble(this.sqlDtRdrServer["PerdTtl_%"]), (GeoPerdasTools.mthdTransformerNoLoadPowerLoss_per(this.sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"])) != 0 && this.intUsaTrafoABNT == 1) ? GeoPerdasTools.mthdTransformerNoLoadPowerLoss_per(this.sqlDtRdrServer["CodFasPrim"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["PotNom_kVA"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh1_kV"])) : Convert.ToDouble(this.sqlDtRdrServer["PerdVz_%"]), this.sqlDtRdrServer["Propr"].ToString(), (this.intAdequarTrafoVazio == 1) ? Convert.ToInt32(this.sqlDtRdrServer["SemCarga"]) : 0);
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";
                    this.addToZipFile(dictLines, @"\TransformadorMTMTMTBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                    
                    this.mthdInputMensageLog("Passou pelos transformadores MT-MT ou MT-BT.");

                    // Loop dos chaves BT
                    this.sqlCommServer = new SqlCommand("SELECT t1.CodChvBT, t1.CodFas, t1.De, t1.Para FROM dbo.[" + this.strBaseSelected + "ChaveBT] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "' AND t1.EstChv = 2", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de chaves BT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandMTSwitch("CBT_" + this.sqlDtRdrServer["CodChvBT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["De"].ToString(), this.sqlDtRdrServer["Para"].ToString());
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\ChavesBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                
                    this.mthdInputMensageLog("Passou pelas chaves BT.");

                    // Loop dos segmentos BT
                    this.sqlCommServer = new SqlCommand("SELECT t1.*, [dbo].[EscolhaPropriedade]([t2].[M1_Propr], [t2].[M2_Propr], [t2].[M3_Propr]) AS [Propr] FROM dbo.[" + this.strBaseSelected + "SegmentoBT] AS t1 LEFT JOIN [dbo].[" + this.strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON ([t1].[CodBase] = [t2].[M1_CodBase] AND [t1].[CodTrafoAtrib] = [t2].[M1_CodTrafo]) WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de segmentos BT";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandLine("SBT_" + this.sqlDtRdrServer["CodSegmBT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl1"].ToString(), this.sqlDtRdrServer["CodPonAcopl2"].ToString(), this.sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["Comp_km"]), this.sqlDtRdrServer["Propr"].ToString(), 0, intNeutralizarRedeTerceiros);
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\SegmentosBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                   
                    this.mthdInputMensageLog("Passou pelos segmentos BT.");

                    // Loop dos ramais
                    this.sqlCommServer = new SqlCommand("SELECT t1.*, [dbo].[EscolhaPropriedade]([t2].[M1_Propr], [t2].[M2_Propr], [t2].[M3_Propr]) AS [Propr] FROM dbo.[" + this.strBaseSelected + "RamalBT] AS t1 LEFT JOIN [dbo].[" + this.strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON ([t1].[CodBase] = [t2].[M1_CodBase] AND [t1].[CodTrafoAtrib] = [t2].[M1_CodTrafo]) WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de ramais";
                    i = 1;
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandLine("RBT_" + this.sqlDtRdrServer["CodRmlBT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl1"].ToString(), this.sqlDtRdrServer["CodPonAcopl2"].ToString(), this.sqlDtRdrServer["CodCond"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["Comp_km"]), this.sqlDtRdrServer["Propr"].ToString(), this.intAdequarRamal, this.intNeutralizarRedeTerceiros);
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";

                    this.addToZipFile(dictLines, @"\RamaisBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                  
                    this.mthdInputMensageLog("Passou pelos ramais.");

                    // Versão usando o bus
                    if (this.config.cbMeterComplete == false)
                    {
                        // Loop dos energymeters de barramento
                        i = 0;
                        this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxMeterBarra AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", this.sqlConnServer);
                        this.sqlCommServer.CommandTimeout = 0;
                        this.strFileData[i] = "! Criação da seção de energymeters de barramento";
                        i++;
                        this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                        while (this.sqlDtRdrServer.Read())
                        {
                            this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandEnergymeter(this.sqlDtRdrServer["Nome"].ToString().TrimEnd(), this.mthdNomeElemento(this.sqlDtRdrServer["Elem"].ToString().TrimEnd()) + this.sqlDtRdrServer["CodNomeElem"].ToString().TrimEnd());
                            i++;
                        }
                        this.sqlDtRdrServer.Close();
                        this.strFileData[i] = "<FIM DOS DADOS>";

                        this.addToZipFile(dictLines, @"\Medidores_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                       
                        this.mthdInputMensageLog("Passou pelos energymeters de barramento.");
                    }

                    // Versão usando a tabela auxiliar de medidores
                    if (this.config.cbMeterComplete == true)
                    {

                        // Loop dos energymeters usando AuxMeter
                        this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxMeterCompleto AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", this.sqlConnServer);
                        this.sqlCommServer.CommandTimeout = 0;
                        this.strFileData[0] = "! Criação da seção de energymeters";
                        i = 1;
                        this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                        while (this.sqlDtRdrServer.Read())
                        {
                            this.strFileData[i] = GeoPerdasTools.mthdWriteLanguageDSSCommandEnergymeter(this.sqlDtRdrServer["Nome"].ToString().TrimEnd(), this.mthdNomeElemento(this.sqlDtRdrServer["Elem"].ToString().TrimEnd()) + this.sqlDtRdrServer["CodNomeElem"].ToString().TrimEnd());
                            i++;
                        }
                        this.sqlDtRdrServer.Close();
                        this.strFileData[i] = "<FIM DOS DADOS>";

                        this.addToZipFile(dictLines, @"\Medidores_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                       
                        this.mthdInputMensageLog("Passou pelos energymeters.");
                    }

                    // Loop das tensões de base
                    strTensoesBase = "Set voltagebases=[";
                    i = 0;
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxVoltageBases AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[0] = "! Criação da seção de tensões de base";
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    while (this.sqlDtRdrServer.Read())
                    {
                        if (i == 0)
                            strTensoesBase += this.sqlDtRdrServer["TnsLnh_kV"].ToString().Replace(",", ".");
                        else
                            strTensoesBase += " " + this.sqlDtRdrServer["TnsLnh_kV"].ToString().Replace(",", ".");
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    strTensoesBase += "]";
                    this.strFileData[1] = strTensoesBase.Replace(",", ".");
                    this.strFileData[2] = "Calcvoltagebases";
                    this.sqlCommServer = new SqlCommand("SELECT t1.* FROM dbo.AuxTensaoBaseBT AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlim = '" + strFeeder + "'", this.sqlConnServer);
                    this.sqlCommServer.CommandTimeout = 0;
                    this.strFileData[3] = "! Criação da seção de reset das tensões de base do BT";
                    this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                    i = 4;
                    while (this.sqlDtRdrServer.Read())
                    {
                        this.strFileData[i] = "!Setkvbase Bus=\"" + this.sqlDtRdrServer["CodPonAcopl"] + "\" kvln=" + this.sqlDtRdrServer["TnsFasBas_kV"].ToString().Replace(",", ".");
                        i++;
                    }
                    this.sqlDtRdrServer.Close();
                    this.strFileData[i] = "<FIM DOS DADOS>";
                   
                    this.addToZipFile(dictLines, @"\Tensoesbase_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                   
                    this.mthdInputMensageLog("Passou pelas tensões de base.");

                    for (int j = 1; j <= 12; j++)
                    {
                        for (int k = 1; k <= 3; k++)
                        {

                            // Loop das cargas MT
                            this.sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + this.strBaseSelected + "AuxCargaMTDem] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                            this.sqlCommServer.CommandTimeout = 0;
                            this.strFileData[0] = "! Criação da seção de cargas MT";
                            i = 1;
                            this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                            while (this.sqlDtRdrServer.Read())
                            {
                                this.strFileData[i] = this.mthdWriteLanguageDSSCommandNewLoad("MT_" + this.sqlDtRdrServer["CodConsMT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, this.sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, this.intAdequarTensaoCargasMT, 0);
                                i++;
                            }
                            this.sqlDtRdrServer.Close();
                            this.strFileData[i] = "<FIM DOS DADOS>";

                            this.addToZipFile(dictLines, @"\CargasMT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                            
                            this.mthdInputMensageLog("Passou pelas cargas MT.");

                            // Loop das cargas BT
                            this.sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV, t2.TnsFasBas_kV FROM dbo.[" + this.strBaseSelected + "AuxCargaBTDem] AS t1 LEFT JOIN dbo.[" + this.strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                            this.sqlCommServer.CommandTimeout = 0;
                            this.strFileData[0] = "! Criação da seção de cargas BT";
                            i = 1;
                            this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                            while (this.sqlDtRdrServer.Read())
                            {
                                this.strFileData[i] = this.mthdWriteLanguageDSSCommandNewLoad("BT_" + this.sqlDtRdrServer["CodConsBT"].ToString(), this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TnsFasBas_kV"]), Convert.ToDouble(this.sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(this.sqlDtRdrServer["DemMaxTrafo_kW"]), this.sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, this.intAdequarTensaoCargasBT, 0);
                                i++;
                            }
                            this.sqlDtRdrServer.Close();
                            this.strFileData[i] = "<FIM DOS DADOS>";

                            this.addToZipFile(dictLines, @"\CargasBT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                            
                            this.mthdInputMensageLog("Passou pelas cargas BT.");

                            if (this.intRealizaCnvrgcPNT == 1)
                            {
                                // Loop das cargas MT Não Técnicas
                                this.sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsMT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, t1.CodAlimAtrib, t1.TnsLnh_kV FROM dbo.[" + this.strBaseSelected + "AuxCargaMTNTDem] AS t1 WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                                this.sqlCommServer.CommandTimeout = 0;
                                this.strFileData[0] = "! Criação da seção de cargas MT Não Técnicas";
                                i = 1;
                                this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                                while (this.sqlDtRdrServer.Read())
                                {
                                    this.strFileData[i] = this.mthdWriteLanguageDSSCommandNewLoad("MT_" + this.sqlDtRdrServer["CodConsMT"].ToString() + "_NT", this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), 0, this.sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 1, this.intAdequarTensaoCargasMT, 0);
                                    i++;
                                }
                                this.sqlDtRdrServer.Close();
                                this.strFileData[i] = "<FIM DOS DADOS>";

                                this.addToZipFile(dictLines, @"\CargasMTNT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");

                                this.mthdInputMensageLog("Passou pelas cargas MT Não Técnicas.");

                                // Loop das cargas BT Não Técnicas
                                this.sqlCommServer = new SqlCommand("SELECT DISTINCT t1.CodBase, t1.CodConsBT, t1.CodFas, t1.CodPonAcopl, t1.TipCrvaCarga, t1.DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW, 0.92 * (ISNULL([M1_PotNom_kVA],0)+ISNULL([M2_PotNom_kVA],0)+ISNULL([M3_PotNom_kVA],0)) AS DemMaxTrafo_kW, t1.CodAlimAtrib, t1.TnsLnh_kV, t2.TnsFasBas_kV FROM dbo.[" + this.strBaseSelected + "AuxCargaBTNTDem] AS t1 LEFT JOIN dbo.[" + this.strBaseSelected + "AuxTrafoMTMTMTBT] AS t2 ON (t1.CodBase = t2.M1_CodBase AND t1.CodTrafoAtrib = t2.M1_CodTrafo) WHERE t1.CodBase = " + this.strBaseSelected + " AND t1.CodAlimAtrib = '" + strFeeder + "'", this.sqlConnServer);
                                this.sqlCommServer.CommandTimeout = 0;
                                this.strFileData[0] = "! Criação da seção de cargas BT Não Técnicas";
                                i = 1;
                                this.sqlDtRdrServer = this.sqlCommServer.ExecuteReader();
                                while (this.sqlDtRdrServer.Read())
                                {
                                    this.strFileData[i] = this.mthdWriteLanguageDSSCommandNewLoad("BT_" + this.sqlDtRdrServer["CodConsBT"].ToString() + "_NT", this.sqlDtRdrServer["CodFas"].ToString(), this.sqlDtRdrServer["CodPonAcopl"].ToString(), Convert.ToDouble(this.sqlDtRdrServer["TnsLnh_kV"]), Convert.ToDouble(this.sqlDtRdrServer["TnsFasBas_kV"]), Convert.ToDouble(this.sqlDtRdrServer["DemMax" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_kW"]), Convert.ToDouble(this.sqlDtRdrServer["DemMaxTrafo_kW"]), this.sqlDtRdrServer["TipCrvaCarga"].ToString() + "_" + GeoPerdasTools.mthdGetTypeDay(k), 2, this.intAdequarTensaoCargasBT, 0);
                                    i++;
                                }
                                this.sqlDtRdrServer.Close();
                                this.strFileData[i] = "<FIM DOS DADOS>";

                                this.addToZipFile(dictLines, @"\CargasBTNT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");
                                
                                this.mthdInputMensageLog("Passou pelas cargas BT Não Técnicas.");
                            }

                            // Loop dos arquivos Master
                            i = 0;
                            this.strFileData[i] = "! Criação da seção do arquivo master"; i++;
                            this.strFileData[i] = "Clear"; i++;
                            this.strFileData[i] = "Redirect 'CircuitoMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'CodCondutor_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'CurvacargaMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'CurvacargaBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'ChavesMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'SegmentosMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'ReguladorMT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'TransformadorMTMTMTBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'SegmentosBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'RamaisBT_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'Medidores_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'CargasMT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Redirect 'CargasBT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            if (this.intRealizaCnvrgcPNT == 1)
                            {
                                this.strFileData[i] = "Redirect 'CargasMTNT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                                this.strFileData[i] = "Redirect 'CargasBTNT_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            }
                            this.strFileData[i] = "Redirect 'Tensoesbase_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss'"; i++;
                            this.strFileData[i] = "Set mode = daily"; i++;
                            this.strFileData[i] = "Set tolerance = 0.0001"; i++;
                            this.strFileData[i] = "Set maxcontroliter = 10"; i++;
                            this.strFileData[i] = "!Set algorithm = newton"; i++;
                            this.strFileData[i] = "!Solve mode = direct"; i++;
                            this.strFileData[i] = "Solve"; i++;
                            this.strFileData[i] = "Export meters"; i++;
                            this.strFileData[i] = "<FIM DOS DADOS>";

                            this.addToZipFile(dictLines, @"\Master_" + GeoPerdasTools.mthdGetTypeDayMonth(j, k) + "_" + this.strBaseSelected + "_" + strFeeder + "_" + this.mthdGetOptionsToString() + ".dss");

                            this.mthdInputMensageLog("Passou pelas criação dos masters.");

                        }

                    }

                    this.mthdInputMensageLog("Término do processo para o alimentador " + strFeeder + ".");

                }

            }
            catch (Exception xcptn)
            {
                this.mthdInputMensageLog("Problema não identificado (" + xcptn.Message.ToString() + ").");
            }            
        }
        return dictLines;

    }

   
}