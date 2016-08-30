import * as React from "react";

import {
  ITable,
  ITableCell
} from "../classes/Interfaces";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  css
} from "office-ui-fabric-react";

import styles from "../ExternalSharingDashboard.module.scss";

class TableCell extends React.Component<ITableCell<any>, {}> {
  private static tableCellClasses: string = css(styles.msTableCellNoWrap, "ms-Table-cell");
  public render(): JSX.Element {
    return (
      <td className={TableCell.tableCellClasses}>{this.props.displayData}</td>
    );
  }
}

export default class Table extends React.Component<ITable, {}> {
  private static tableClasses: string = css("ms-Table");
  private static tableRowClasses: string = css("ms-Table-row");

  public render(): JSX.Element {
      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <div className={styles.msTableOverflow}>
              <table className={Table.tableClasses}>
                <thead>
                  <tr className={Table.tableRowClasses}>
                    {this.props.columns.cells.map(c => {
                        return (
                          <TableCell {...c} />
                        );
                      })}
                  </tr>
                </thead>
                <tbody>
                    {this.props.rows.map(r => {
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
            </div>
        </FocusZone>
      );
  }
}
