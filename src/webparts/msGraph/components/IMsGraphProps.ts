import { MSGraphClient } from "@microsoft/sp-http";

export interface IMsGraphProps {
  description: string;
  graphClient: MSGraphClient;
}

export interface IChart {
  chartData: {};
}

