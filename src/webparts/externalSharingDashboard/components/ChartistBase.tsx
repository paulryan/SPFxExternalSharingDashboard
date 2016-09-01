import * as Chartist from "chartist";
import * as React from "react";

import "../DocumentDashboard.module.css";

import {
  IChart,
  IChartItem
} from "../classes/Interfaces";

export abstract class ChartistBase extends React.Component<IChart, IChart> {

  protected responsiveOptions: Chartist.IResponsiveOptionTuple<Chartist.IPieChartOptions>[] = [
    ["screen and (max-width: 1024px)", {
      labelOffset: 0,
      chartPadding: 5
    }]
  ];

  public render(): JSX.Element {
      return (
        <div id="chartist" className="ct-chart ct-golden-section"></div>
      );
  }

  public componentDidMount(): void {
    this.renderChart();
  }

  public abstract renderChart(): void;

  protected getChartistData(): Chartist.IChartistData {
    // Create a object of chart items
    const chartItemDatas: string[] = [];
    const chartItemsDict: any = {};

    this.props.items.forEach(dataPoint => {
      const dataPointFromDict: IChartItem = chartItemsDict[dataPoint.data];
      if (dataPointFromDict) {
        dataPointFromDict.weight++;
      }
      else {
        chartItemsDict[dataPoint.data] = dataPoint;
        chartItemDatas.push(dataPoint.data);
      }
    });

    chartItemDatas.sort();

    const labels: string[] = chartItemDatas.map<string>(data => (chartItemsDict[data] as IChartItem).label);
    const dataSeries: number[] = chartItemDatas.map<number>(data => (chartItemsDict[data] as IChartItem).weight);
    const data: Chartist.IChartistData = {
      labels: labels,
      series: dataSeries
    };
    return data;
  }

  protected reduceSumIChartistSeriesData(prev: Chartist.IChartistSeriesData, curr: Chartist.IChartistSeriesData): number {
    return this.reduceSumNumber(prev.value, curr.value);
  }

  protected reduceSumNumber(prev: number, curr: number): number {
    return prev + curr;
  }
}
