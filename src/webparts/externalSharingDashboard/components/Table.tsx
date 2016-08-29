import * as React from "react";

import {
  ISecurableObject,
  ITableProps,
  SecurableObjectType
} from "../classes/Interfaces";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  css
} from "office-ui-fabric-react";

import styles from "../ExternalSharingDashboard.module.scss";

class TableRow extends React.Component<ISecurableObject, {}> {
  private static tableRowClasses: string = css("ms-Table-row");
  private static tableCellClasses: string = css(styles.msTableCellNoWrap, "ms-Table-cell");

  public render(): JSX.Element {
    return (
      <tr className={TableRow.tableRowClasses}>
        <td className={TableRow.tableCellClasses}>{this.props.CrawlTime}</td>
        <td className={TableRow.tableCellClasses}>{this.props.Title}</td>
        <td className={TableRow.tableCellClasses}>{this.props.SharedWith}</td>
        <td className={TableRow.tableCellClasses}>{this.props.SharedBy}</td>
        <td className={TableRow.tableCellClasses}>{this.props.SiteTitle}</td>
        <td className={TableRow.tableCellClasses}>{this.props.FileExtension}</td>
      </tr>
    );
  }
}

export default class Table extends React.Component<ITableProps, {}> {
  private static tableClasses: string = css("ms-Table");

  public render(): JSX.Element {
      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <div className={styles.msTableOverflow}>
              <table className={Table.tableClasses}>
                <thead>
                  <TableRow
                    key="headerRow"
                    Type={SecurableObjectType.Document}
                    Title="Title"
                    FileExtension="File Extension"
                    LastModifiedTime="Last Modified Time"
                    SiteTitle="Site Title"
                    SiteID="Site ID"
                    SharedBy={[]}
                    SharedWith={[]}
                    CrawlTime="Crawl Time"
                    URL="URL" />
                </thead>
                <tbody>
                  {this.props.items.map(c => {
                    return (
                      <TableRow {...c} />
                    );
                  })}
                </tbody>
              </table>
            </div>
        </FocusZone>
      );
  }
}
