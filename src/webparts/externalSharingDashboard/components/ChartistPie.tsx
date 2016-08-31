import * as Chartist from "chartist";
import * as React from "react";

import "../ExternalSharingDashboard.module.css";

import {
  IChart,
  IChartItem
} from "../classes/Interfaces";

export default class ChartistPie extends React.Component<IChart, IChart> {

  private columnCount: number = 0;
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
    if (this.props.rows.length > 0) {
      this.columnCount = this.props.rows[0].cells.length;
      if (this.props.columnIndexToGroupUpon >= 0 && this.props.columnIndexToGroupUpon < this.columnCount) {
        // set state
        this.setState(this.props);
      }
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
    if (currentState && currentState.rows && currentState.columnIndexToGroupUpon) {
      // Create a object of chart items
      const chartItemLabels: string[] = [];
      const chartItemsDict: any = {};

      currentState.rows.forEach(r => {
        const d: string = r.cells[currentState.columnIndexToGroupUpon].displayData;
        let chartItem: IChartItem = chartItemsDict[d];
        if (chartItem) {
          chartItem.value++;
        }
        else {
          chartItem = {
            label: d,
            value: 1
          };
          chartItemsDict[d] = chartItem;
          chartItemLabels.push(chartItem.label);
        }
      });

      const dataSeries: number[] = chartItemLabels.map<number>(l => (chartItemsDict[l] as IChartItem).value);
      const data: Chartist.IChartistData = {
        labels: chartItemLabels,
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
