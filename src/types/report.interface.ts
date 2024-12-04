export interface IRep {
  badDataCount: number;
  goodDataCount: number;
  gov: any[];
  mange: any[];
  center: any[];
  classes: any[];
  cls: {
    [key: string]: string[];
  };
}
