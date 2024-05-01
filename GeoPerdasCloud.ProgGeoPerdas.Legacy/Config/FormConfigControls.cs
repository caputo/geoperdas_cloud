using System;

namespace GeoPerdasCloud.ProgGeoPerdas.Legacy.Config
{
    public class FormConfigControls
    {
        public int cbModel { get; set; }//Modelo de cargas 1,2
        public string tbServer { get; set; }
        public string tbDataBase { get; set; }
        public string tbUser { get; set; }
        public string tbPassword { get; set; }
        public string tbObjectFeeder { get; set; }
        public string tbObjectBase { get; set; }
        public string rtbLog { get; set; }
        public string tbPathFile { get; set; }
        public bool cbReinicializeNT { get; set; }
        public string tbVPUMin { get; set; }
        public bool cbMeterComplete { get; set; }
        public string rtbInfo { get; set; }
        public string rtbAdmin { get; set; }
        public string tbCommand { get; set; }
        public int intRealizaCnvrgcPNT { get; set; }
        public int intUsaTrafoABNT { get; set; }
        public int intAdequarTensaoCargasMT { get; set; }
        public int intAdequarTensaoCargasBT { get; set; }
        public int intAdequarTensaoSuperior { get; set; }
        public int intAdequarRamal { get; set; }
        public int intAdequarTapTrafo { get; set; }
        public int intAdequarPotenciaCarga { get; set; }
        public int dblVPUMin { get; set; }
        public int intAdequarTrafoVazio { get; set; }
        public int intAdequarModeloCarga { get; set; }
        public int intNeutralizarTrafoTerceiros { get; set; }
        public int intNeutralizarRedeTerceiros { get; set; }
        
        public IDictionary<string, bool> clbOption { get; set; }

        public bool transformadoresABNT { get; set; }
        public bool adequaMT { get; set; }
        public bool adequaBT { get; set; }
        public bool limiteTensaoBarra { get; set; }
        public bool limiteRamal { get; set; }
        public bool tapTransformadores { get; set; }
        public bool limiteCargaBT { get; set; }
        public bool neutralizaTrafo { get; set; }
        public bool netralizaRede { get; set; }
        public bool trafoVazio { get; set; }

        public FormConfigControls Clone() {
            return (FormConfigControls)this.MemberwiseClone();
        }
    }
}