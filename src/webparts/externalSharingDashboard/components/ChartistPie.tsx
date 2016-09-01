import * as Chartist from "chartist";

import "../DocumentDashboard.module.css";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistPie extends ChartistBase {

  public renderChart(): void {

    const data: Chartist.IChartistData = this.getChartistData();
    const dataSeries: number[] = data.series as number[];
    const seriesTotal: number = dataSeries.reduce(this.reduceSumNumber);

    const options: Chartist.IPieChartOptions = {
      chartPadding: {
        top: 30,
        right: 30,
        bottom: 30,
        left: 30
      },
      labelOffset: 110,
      labelDirection: "explode",
      labelInterpolationFnc: (label: string, index: number): string => {
        const valueAsNumber: number = data.series[index] as number;
        const valueAsPercentage: number = Math.round(valueAsNumber / seriesTotal * 1000) / 10;
        return label + " (" + valueAsPercentage + "%)";
      }
    };

    let pie: any = new Chartist.Pie("#chartist", data, options, this.responsiveOptions);
    // tslint ignore..
    pie = pie;
  }
}
