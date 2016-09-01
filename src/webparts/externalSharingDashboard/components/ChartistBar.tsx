import * as Chartist from "chartist";

import "../DocumentDashboard.module.css";

import {
  ChartistBase
} from "./ChartistBase";

export default class ChartistBar extends ChartistBase {

  public renderChart(): void {

    const data: Chartist.IChartistData = this.getChartistData();

    const options: Chartist.IBarChartOptions = {
      axisY: {
        onlyInteger: true
      },
      chartPadding: {
        top: 30,
        right: 30,
        bottom: 30,
        left: 30
      }
    };

    // Bar graphs take an array of series as they support many bars
    data.series = [ data.series ];

    let bar: any = new Chartist.Bar("#chartist", data, options, this.responsiveOptions);
    // tslint ignore..
    bar = bar;
  }
}
