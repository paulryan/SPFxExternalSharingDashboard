import * as React from "react";

import {
  ITableProps
} from "../classes/Interfaces";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  css

} from "office-ui-fabric-react";

import styles from "../ExternalSharingDashboard.module.scss";

export default class Table extends React.Component<ITableProps, {}> {
  public render(): JSX.Element {
      const tableClasses: string = css("ms-Table");
      const tableRowClasses: string = css("ms-Table-row");
      const tableCellClasses: string = css(styles.msTableCellNoWrap, "ms-Table-cell");
      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <div className={styles.msTableOverflow}>
              <table className={tableClasses}>
                  <tr className={tableRowClasses}>
                    <td className={tableCellClasses}>Type</td>
                    <td className={tableCellClasses}>Title</td>
                    <td className={tableCellClasses}>Modified</td>
                    <td className={tableCellClasses}>Shared With</td>
                    <td className={tableCellClasses}>Shared By</td>
                  </tr>
                <tbody>
                  {this.props.items.map(c => {
                    return (
                      <tr className={tableRowClasses}>
                        <td className={tableCellClasses}>{c.Type}</td>
                        <td className={tableCellClasses}>{c.Title}</td>
                        <td className={tableCellClasses}>{c.LastModifiedTime}</td>
                        <td className={tableCellClasses}>{c.SharedWith}</td>
                        <td className={tableCellClasses}>{c.SharedBy}</td>
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
