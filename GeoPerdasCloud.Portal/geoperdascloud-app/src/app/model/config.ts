export class GeoPerdasConfig{
    public tbServer!: string;    
    public tbDatabase!: string;    
    public tbUser!: string;    
    public tbPassword!: string;    
    public tbObjectBase!: string;    
    public tbObjectFeeder!: string;
    public cbReinicializeNT: boolean = false;
    public cbMeterComplete: boolean = false;
    public tbVPUMin!:Number;
    public cbModel!:Number;
    public convergiaPNT:boolean = false;
    public transformadoresABNT:boolean = false;
    public adequaMT:boolean = false;
    public adequaBT:boolean = false;
    public limiteTensaoBarra:boolean = false;
    public limiteRamal:boolean = false;
    public tapTransformadores:boolean = false;
    public limiteCargaBT:boolean = false;
    public neutralizaTrafo:boolean = false;
    public netralizaRede:boolean = false;
    public trafoVazio:boolean = false;

}