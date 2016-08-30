import * as React from "react";

import {
  ControlMode,
  GetDisplayTermForEnumMode,
  GetDisplayTermForEnumSPScope
} from "../classes/Enums";

import {
  IExternalSharingDashboardProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectProperty,
  ITableCell,
  ITableRow
} from "../classes/Interfaces";

import {
  Logger
} from "../classes/Logger";

import {
  Label
} from "office-ui-fabric-react";

import Table from "./Table";

interface ITableProps {
  items: ISecurableObject[];
}

interface ITableRowProps {
  item: ISecurableObject;
}

export default class ExternalSharingDashboard extends React.Component<IExternalSharingDashboardProps, IGetExtContentFuncResponse> {
  private log: Logger;
  private isUpdateStateInProgress: boolean = false;

  // Lifecycle methods are called as follows:
  // componentWillMount     (set state to Loading)
  // render                 (loading)
  // componentDidMount      (fetch ext content)
  // shouldComponentUpdate  (on response received)
  // render                 (content)
  // componentDidUpdate     (ignored as request in progress..?)

  constructor() {
    super();
    this.log = new Logger("ExternalSharingDashboard");
  }

  public componentWillMount(): void {
    this.log.logInfo("componentWillMount");
    this._setStateToLoading();
  }

  public componentDidMount(): void {
    this.log.logInfo("componentDidMount");
    this._updateState();
  }

  public componentWillReceiveProps(): void {
    this.log.logInfo("componentWillReceiveProps");
    this._setStateToLoading();
  }

  public shouldComponentUpdate(nextProps: IExternalSharingDashboardProps): boolean {
    this.log.logInfo("shouldComponentUpdate");
    return !this.state || nextProps.contentProps.mode !== this.state.mode
      || nextProps.contentProps.scope !== this.state.scope;
  }

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    this._updateState();
  }

  public render(): JSX.Element {
    this.log.logInfo("render");
    const headerControls: JSX.Element = (
      <div>
        <div className="ms-font-xxl">External Sharing Dashboard</div>
        <div className="ms-font-l">{GetDisplayTermForEnumMode(this.state.mode) + " " + GetDisplayTermForEnumSPScope(this.state.scope).toLowerCase()}</div>
      </div>
    );

    if (this.state && this.state.controlMode === ControlMode.Loading) {
      return (
        <div>
          {headerControls}
          <div className="ms-font-l">{this.state.message}</div>
        </div>
      );
      //<Spinner type={ SpinnerType.large } label={this.state.message} />
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      return (
        <div>
          {headerControls}
          <Label>{this.state.message}</Label>
        </div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {

      // TODO : In cases with lots of data it will not be okay to process all data
      // upfront - only the current page should be processed?
      const columnWithHref: string = "title";
      const columns: ITableCell<string>[] = [
        { sortableData: "title", displayData: "Title", href: null, key: "headerCellTitle"},
        { sortableData: "sharedWith", displayData: "Shared With", href: null, key: "headerCellSharedWith"},
        { sortableData: "sharedBy", displayData: "Shared By", href: null, key: "headerCellSharedBy"},
        { sortableData: "siteTitle", displayData: "Site Title", href: null, key: "headerCellSiteTitle"},
        { sortableData: "crawlTime", displayData: "Accurate as of", href: null, key: "headerCellCrawlTime"}
      ];

      const rows: ITableRow[] = [];
      this.state.extContent.forEach((securableObj) => {
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

      return (
        <div>
          {headerControls}
          <Table columns={{cells:columns, key:"headerRow" }} rows={rows} pageSize={10} pageStartIndex={0} currentSort={-1} currentSortDescending={true} />
        </div>
      );
    }
    else if (this.state && this.state.controlMode) {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
      return (<div className="ms-font-l">Error!</div>);
    }
    else {
      this.log.logError(`State is undefined`);
      //this.log.logInfo(`State is undefined`);
      // This will occur is the state is not set in componentWillMount
      return (
        <div className="ms-font-l">Error!</div>
      );
    }
  }

  private _updateState(): void {
    if (!this.isUpdateStateInProgress) {
      this.isUpdateStateInProgress = true;
      this.props.store.getExternalContent()
      .then((r) => {
        this.setState(r);
        this.isUpdateStateInProgress = false;
      });
    }
    else {
      this.log.logInfo("update state ignored as request is already in progress");
    }
  }

  private _setStateToLoading(): void {
    this.setState({
      extContent: [],
      controlMode: ControlMode.Loading,
      message: "Working on it...",
      mode: -1,
      scope: -1
    });
  }
}
