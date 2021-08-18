export class ColumnAPI {
  TradingPeriod: string;
  Market: string;
  B: number;
  BVol: number;
  A: number;
  AVol: number;
  IBAdd: number;
  IBVol: number;
  IAAdd: number;
  IAVol: number;
  IWap: number;

  constructor(
    TradingPeriod: string,
    Market: string,
    B: number,
    BVol: number,
    A: number,
    AVol: number,
    IBAdd: number,
    IBVol: number,
    IAAdd: number,
    IAVol: number,
    IWap: number
  ) {
    this.TradingPeriod = TradingPeriod;
    this.Market = Market;
    this.B = B;
    this.BVol = BVol;
    this.A = A;
    this.AVol = AVol;
    this.IBAdd = IBAdd;
    this.IBVol = IBVol;
    this.IAAdd = IAAdd;
    this.IAVol = IAVol;
    this.IWap = IWap;
  }
}
