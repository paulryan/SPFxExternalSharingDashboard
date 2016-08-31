import * as React from "react";

import {
  ControlMode,
  DisplayType,
  GetDisplayTermForEnumMode,
  GetDisplayTermForEnumSPScope
} from "../classes/Enums";

import {
  IChart,
  IChartItem,
  IDocumentDashboardProps,
  IDocumentDashboardState,
  ISecurableObject,
  ISecurableObjectProperty,
  ITable,
  ITableCell,
  ITableRow
} from "../classes/Interfaces";

import {
  Logger
} from "../classes/Logger";

import {
  Label
} from "office-ui-fabric-react";

import ChartistPie from "./ChartistPie";
import Table from "./Table";

interface ITableProps {
  items: ISecurableObject[];
}

interface ITableRowProps {
  item: ISecurableObject;
}

export default class DocumentDashboard extends React.Component<IDocumentDashboardProps, IDocumentDashboardState> {
  private log: Logger;
  private isUpdateStateInProgress: boolean = false;
  private hasContentBeenFetched: boolean = false;

  // Lifecycle methods are called as follows:
  // componentWillMount     (set state to Loading)
  // render                 (loading)
  // componentDidMount      (fetch ext content)
  // shouldComponentUpdate  (on response received)
  // render                 (content)
  // componentDidUpdate     (ignored as request in progress..?)

  constructor() {
    super();
    this.log = new Logger("DocumentDashboard");
  }

  public componentWillMount(): void {
    this.log.logInfo("componentWillMount");
    this.setStateWrapper([], ControlMode.Loading, "Working on it...");
  }

  public componentDidMount(): void {
    this.log.logInfo("componentDidMount");
    this.updateState();
  }

  public componentWillReceiveProps(): void {
    this.log.logInfo("componentWillReceiveProps");
    this.setStateWrapper(this.state.results, ControlMode.Loading, "Working on it...");
  }

  public shouldComponentUpdate(nextProps: IDocumentDashboardProps, nextState: IDocumentDashboardState): boolean {
    this.log.logInfo("shouldComponentUpdate");
    return !this.state
      || this.state.controlMode !== nextState.controlMode
      || this.state.mode !== nextState.mode
      || this.state.scope !== nextState.scope
      || this.state.displayType !== nextState.displayType;
  }

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    this.updateState();
  }

  public render(): JSX.Element {
    // Reusable components
    const headerControls: JSX.Element = (
      <div>
        <div className="ms-font-xxl">Document Dashboard</div>
        <div className="ms-font-l">{GetDisplayTermForEnumMode(this.state.mode) + " " + GetDisplayTermForEnumSPScope(this.state.scope).toLowerCase()}</div>
      </div>
    );

    // Render according to the control mode
    if (this.state && this.state.controlMode === ControlMode.Loading) {
      this.log.logInfo("render (Loading)");
      return (
        <div>
          {headerControls}
          <div className="ms-font-l">{this.state.message}</div>
        </div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      this.log.logInfo("render (Message)");
      return (
        <div>
          {headerControls}
          <Label>{this.state.message}</Label>
        </div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {
      this.log.logInfo("render (Content)");

      // Select the appropriate comnponent
      let component: JSX.Element = null;
      if (this.state.displayType === DisplayType.Table) {
        const params: ITable = this.getStateAsITable();
        component = (
          <Table {...params} />
        );
      }
      else if (this.state.displayType === DisplayType.BySite || this.state.displayType === DisplayType.ByUser) {
        const params: IChart = this.getStateAsIChart(this.state.displayType);
        component = (
          <ChartistPie {...params} />
        );
      }
      else if (this.state.displayType === DisplayType.OverTime) {
        const params: IChart = this.getStateAsIChart(this.state.displayType);
        component = (
          <ChartistPie {...params} />
        );
      }
      else {
        this.log.logError("Unsupported display type: " + this.state.displayType);
        return null;
      }

      return (
        <div>
          {headerControls}
          {component}
        </div>
      );
    }
    else if (this.state && this.state.controlMode) {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
      return (
        <div className="ms-font-l">Error!</div>
      );
    }
    else {
      this.log.logError(`State is undefined`);
      return (
        <div className="ms-font-l">Error!</div>
      );
    }
  }

  private shouldFetchContent(): boolean {
    return !this.state || !this.hasContentBeenFetched
      || this.props.mode !== this.state.mode
      || this.props.scope !== this.state.scope;
  }

  private updateState(): void {
    if (!this.isUpdateStateInProgress) {
      this.isUpdateStateInProgress = true;
      if (this.shouldFetchContent()) {
        this.props.store.getContent()
        .then((r) => {
          const controlMode: ControlMode = r.isError || r.results.length < 1 ? ControlMode.Message : ControlMode.Content;
          this.setStateWrapper(r.results, controlMode, r.message);
          this.hasContentBeenFetched = true;
          this.isUpdateStateInProgress = false;
        })
        .catch((e) => {
          this.log.logError("Failed to get content", e.message ? e.message : e.toString());
          this.setStateWrapper(this.state.results, ControlMode.Message, "Failed to get content");
          this.isUpdateStateInProgress = false;
        });
      }
      else {
        this.log.logInfo("New content has not been fetched as only the display mode has changed");
        // This sets the mode, scope, display mode as per props without changing the data
        this.setStateWrapper(this.state.results, ControlMode.Content, "Using cached content");
        this.isUpdateStateInProgress = false;
      }
    }
    else {
      this.log.logInfo("Update state ignored as request is already in progress");
    }
  }

  private setStateWrapper(results: ISecurableObject[], controlMode: ControlMode, message: string): void {
    this.setState({
      results: results,
      controlMode: controlMode,
      message: message,
      mode: this.props.mode,
      scope: this.props.scope,
      displayType: this.props.displayType
    });
  }

  // columns={{cells:columns, key:"headerRow" }} rows={rows} columnIndexToGroupUpon={3}
  private getStateAsIChart(displayType: DisplayType): IChart {
    const dataPoints: IChartItem[] = [];
    this.state.results.forEach((securableObj) => {
      if (displayType === DisplayType.ByUser) {
        securableObj.sharedBy.data.forEach((d) => {
          dataPoints.push({
            label: d,
            value: 1
          });
        });

        // Add data points for modified by
      }
      else if (displayType === DisplayType.BySite) {
        dataPoints.push({
          label: securableObj.siteTitle.displayValue,
          value: 1
        });
      }
      else if (displayType === DisplayType.OverTime) {
        if (securableObj.lastModifiedTime.data) {
          dataPoints.push({
            label: securableObj.lastModifiedTime.displayValue,
            value: 1
          });
        }
      }
    });
    return {
      items: dataPoints,
      columnIndexToGroupUpon: 0 // As there is only a single column in the data we return
    };
  }

  private getStateAsITable(): ITable {
    // TODO : In cases with lots of data it will not be okay to process all data
    // upfront - only the current page should be processed?
    const columnWithHref: string = "title";
    const columns: ITableCell<string>[] = [
      { sortableData: "title", displayData: "Title", href: null, key: "headerCellTitle"},
      { sortableData: "lastModifiedTime", displayData: "Modified", href: null, key: "headerCellModified"},
      { sortableData: "sharedWith", displayData: "Shared With", href: null, key: "headerCellSharedWith"},
      { sortableData: "sharedBy", displayData: "Shared By", href: null, key: "headerCellSharedBy"},
      { sortableData: "siteTitle", displayData: "Site Title", href: null, key: "headerCellSiteTitle"},
      { sortableData: "crawlTime", displayData: "Accurate as of", href: null, key: "headerCellCrawlTime"}
    ];

    const rows: ITableRow[] = [];
    this.state.results.forEach((securableObj) => {
      const newRow: ITableRow = { cells: [], key: securableObj.key};
      columns.forEach((columnName) => {
        const cellSortableData: ISecurableObjectProperty<any> = securableObj[columnName.sortableData];
        if (cellSortableData) {
          const href: string = (columnName.sortableData === columnWithHref ? securableObj.url.data : null);
          newRow.cells.push({
            sortableData: cellSortableData.data,
            displayData: cellSortableData.displayValue,
            href: href,
            key: columnName.sortableData
          });
        }
        else {
          this.log.logError("Column value not present on row: " + columnName.sortableData);
          // Still add a cell to so that the rows do go out of line
          newRow.cells.push({
            sortableData: "?",
            displayData: "",
            href: null,
            key: securableObj.key + columnName.sortableData
          });
        }
      });
      rows.push(newRow);
    });

    return {
      columns: { cells: columns, key: "headerRow" },
      rows: rows,
      currentSort: -1,
      currentSortDescending: false,
      pageSize: 10,
      pageStartIndex: 0
    };
  }
}
