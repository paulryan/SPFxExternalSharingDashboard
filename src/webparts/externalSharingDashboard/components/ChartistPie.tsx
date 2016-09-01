import * as Chartist from "chartist";
import * as React from "react";

import "../DocumentDashboard.module.css";

import {
  IChart,
  IChartItem
} from "../classes/Interfaces";

export default class ChartistPie extends React.Component<IChart, IChart> {

  //private columnCount: number = 0;
  private responsiveOptions: Chartist.IResponsiveOptionTuple<Chartist.IPieChartOptions>[] = [
    ["screen and (max-width: 1024px)", {
      labelOffset: 0,
      chartPadding: 5
    }]
  ];
  public render(): JSX.Element {
      return (
        <div id="chartistPie" className="ct-chart ct-golden-section"></div>
      );
  }

  public componentWillMount(): void {
    // Calculate constants
    if (this.props.items.length > 0) {
      // set state
      this.setState(this.props);
    }
  }

  public componentDidMount(): void {
    // Init render of chart
    this.renderChart();
  }

  public componentDidUpdate(): void {
    // Non init render of char
    this.renderChart();
  }

  private renderChart(): void {
    const currentState: IChart = this.state;
    if (currentState && currentState.items) {
      // Create a object of chart items
      const chartItemDatas: string[] = [];
      const chartItemsDict: any = {};

      currentState.items.forEach(dataPoint => {
        const dataPointFromDict: IChartItem = chartItemsDict[dataPoint.data];
        if (dataPointFromDict) {
          dataPointFromDict.weight++;
        }
        else {
          chartItemsDict[dataPoint.data] = dataPoint;
          chartItemDatas.push(dataPoint.data);
        }
      });

      const labels: string[] = chartItemDatas.map<string>(data => (chartItemsDict[data] as IChartItem).label);
      const dataSeries: number[] = chartItemDatas.map<number>(data => (chartItemsDict[data] as IChartItem).weight);
      const data: Chartist.IChartistData = {
        labels: labels,
        series: dataSeries
      };

      const seriesTotal: number = dataSeries.reduce(this.reduceSum);

      const options: Chartist.IPieChartOptions = {
        chartPadding: 30,
        labelOffset: 110,
        labelDirection: "explode",
        labelInterpolationFnc: (label: string, index: number): string => {
          const valueAsNumber: number = dataSeries[index];
          const valueAsPercentage: number = Math.round(valueAsNumber / seriesTotal * 1000) / 10;
          return label + " (" + valueAsPercentage + "%)";
        }
      };

      let pie: any = new Chartist.Pie("#chartistPie", data, options, this.responsiveOptions);
      // tslint ignore..
      pie = pie;
    }
  }

  private reduceSum (prev: number, curr: number): number {
    return prev + curr;
  }
}
