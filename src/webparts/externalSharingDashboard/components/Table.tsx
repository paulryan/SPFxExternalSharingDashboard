import * as React from "react";

import {
  ITable,
  ITableCell,
  ITableRow
} from "../classes/Interfaces";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  css
} from "office-ui-fabric-react";

import styles from "../DocumentDashboard.module.scss";

export default class Table extends React.Component<ITable, ITable> {
  private static tableClasses: string = css("ms-Table");
  private static tableHeadClasses: string = css("ms-bgColor-neutralLight");
  private static tableBodyClasses: string = css("ms-bgColor-white");
  private static tableRowClasses: string = css("ms-Table-row");
  private static tablePagerClasses: string = css(styles.msTablePager, "ms-font-l");

  private maxPageStartIndex: number = 0;
  private minPageStartIndex: number = 0;
  private pageCount: number = 1;
  private columnCount: number = 0;

  constructor() {
    super();
    this.nextPage = this.nextPage.bind(this);
    this.prevPage = this.prevPage.bind(this);
    this.sortOnColumn = this.sortOnColumn.bind(this);
  }

  public componentWillMount(): void {
    // Calcuate constants
    const rowCount: number = this.props.rows.length;
    this.maxPageStartIndex = rowCount - 1; //rowCount - this.props.pageSize;
    if (this.maxPageStartIndex < this.minPageStartIndex) {
      this.maxPageStartIndex = this.minPageStartIndex;
    }
    this.pageCount = Math.ceil(rowCount / this.props.pageSize);
    if (rowCount > 0) {
      this.columnCount = this.props.rows[0].cells.length;
    }

    // Ensure props are in range
    // if (this.props.pageSize < 1) {
    //   this.props.pageSize = 1;
    // }
    // if (this.props.pageStartIndex < this.minPageStartIndex) {
    //   this.props.pageSize = this.minPageStartIndex;
    // }
    // if (this.props.pageStartIndex > this.maxPageStartIndex) {
    //   this.props.pageSize = this.maxPageStartIndex;
    // }

    // Support initial sort
    if (this.props.currentSort >= 0 && this.props.currentSort < this.columnCount) {
      this.sortOnColumnInternal(this.props);
    }

    this.setState(this.props);
  }

  public render(): JSX.Element {
      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <div className={styles.msTableOverflow}>
              <table className={Table.tableClasses}>
                <thead className={Table.tableHeadClasses}>
                  <tr className={Table.tableRowClasses}>
                    {this.state.columns.cells.map((c, i) => {
                        return (
                          <TableCell {...c} onClick={() => this.sortOnColumn(i)} />
                        );
                      })}
                  </tr>
                </thead>
                <tbody className={Table.tableBodyClasses}>
                    {this.state.rows.slice(this.state.pageStartIndex, this.state.pageStartIndex + this.state.pageSize).map(r => {
                      return (
                        <tr key={r.key} className={Table.tableRowClasses}>
                          {r.cells.map(c => {
                            return (
                              <TableCell {...c} />
                            );
                          })}
                        </tr>
                      );
                    })}
                </tbody>
              </table>
              <div className={Table.tablePagerClasses}>
                <a href="#" onClick={this.prevPage}>
                  <i className="ms-Icon ms-Icon--triangleLeft" aria-action="Previous page of results"></i>
                </a>
                <span>{Math.round(this.state.pageStartIndex / this.state.pageSize) + 1} of {this.pageCount}</span>
                <a href="#" onClick={this.nextPage}>
                  <i className="ms-Icon ms-Icon--triangleRight" aria-action="Next page of results"></i>
                </a>
              </div>
            </div>
        </FocusZone>
      );
  }

  private nextPage (): void {
    const newPageStartIndex: number = this.state.pageStartIndex + this.state.pageSize;
    if (newPageStartIndex > this.maxPageStartIndex) {
      // already on the last page
    }
    else {
      this.state.pageStartIndex = newPageStartIndex;
      this.setState(this.state);
    }
  }

  private prevPage (): void {
    let newPageStartIndex: number = this.state.pageStartIndex - this.state.pageSize;
    if (newPageStartIndex < this.minPageStartIndex) {
      // Ensure on the first page
      newPageStartIndex = this.minPageStartIndex;
    }
    this.state.pageStartIndex = newPageStartIndex;
    this.setState(this.state);
  }

  private sortOnColumn (columnIndex: number): void {
      if (columnIndex >= 0 && columnIndex < this.columnCount) {
        // if the column is already sorted, sort in opposite direction, else sort in constant direction
        this.state.currentSortDescending = this.state.currentSort === columnIndex ? !this.state.currentSortDescending : false;
        this.state.currentSort = columnIndex;

        // do the sorting
        this.sortOnColumnInternal(this.state);

        // trigger the update
        this.setState(this.state);
      }
  }

  private sortOnColumnInternal (data: ITable): void {
    if (data.currentSort >= 0 && data.currentSort < this.columnCount) {
      data.rows.sort((rowA, rowB) => {
        const cmpr: number = this.compareTableRow(data.currentSort, rowA, rowB);
        return (data.currentSortDescending ? cmpr * -1 : cmpr);
      });
    }
  }

  private compareTableRow (columnIndex: number, rowA: ITableRow, rowB: ITableRow): number {
    const cellA: ITableCell<any> = rowA.cells[columnIndex];
    const cellB: ITableCell<any> = rowB.cells[columnIndex];
    const cellDataA: any = cellA.sortableData;
    const cellDataB: any = cellB.sortableData;

    let compareValue: number = 0;
    if (typeof cellDataA === "string" && typeof cellDataB === "string") {
      compareValue = cellDataA.localeCompare(cellDataB);
    }
    else if (cellDataA instanceof Date && cellDataB instanceof Date) {
      compareValue = (cellDataA > cellDataB) ? 1
                      : (cellDataA < cellDataB) ? -1 : 0;
    }
    else if (cellDataA instanceof Array && cellDataB instanceof Array) {
      // sort on the display values... not best solution
      compareValue = cellA.displayData.localeCompare(cellB.displayData);
    }
    return compareValue;
  }
}

class TableCell extends React.Component<ITableCell<any>, {}> {
  private static tableCellClasses: string = css(styles.msTableCellNoWrap, "ms-Table-cell");
  private static tableCellHyperlinkClasses: string = css("ms-Link");
  public render(): JSX.Element {
    if (this.props.href) {
      return (
        <td className={TableCell.tableCellClasses}>
          <a className={TableCell.tableCellHyperlinkClasses} href={this.props.href} target="_blank">
            {this.props.displayData}
          </a>
        </td>
      );
    }
    else if (this.props.onClick) {
      // onClick={e => _self.handleClick(cellVal)}
      return (
        <td className={TableCell.tableCellClasses} onClick={this.props.onClick}>
          {this.props.displayData}
        </td>
      );
    }
    else {
      return (
        <td className={TableCell.tableCellClasses}>
          {this.props.displayData}
        </td>
      );
    }
  }
}