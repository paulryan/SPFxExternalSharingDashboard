import {
  IWebPartContext
} from "@microsoft/sp-client-preview";

import {
  ControlMode,
  DisplayType,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Enums";

export interface IExtContentFetcherProps {
  context: IWebPartContext;
  scope: SPScope;
  mode: Mode;
  managedProperyName: string;
  noResultsString: string;
}

export interface IGetExtContentFuncResponse {
  extContent: ISecurableObject[];
  controlMode: ControlMode;
  message: string;
  scope: SPScope;
  mode: Mode;
}

export interface ISecurableObject {
  // That match managed property names
  title: ISecurableObjectProperty<string>;
  fileExtension: ISecurableObjectProperty<string>;
  lastModifiedTime: ISecurableObjectProperty<Date>;
  siteTitle: ISecurableObjectProperty<string>;
  siteID: ISecurableObjectProperty<string>;
  crawlTime: ISecurableObjectProperty<Date>;

  // That require mapping/transforming from managed property
  url: ISecurableObjectProperty<string>;
  type: ISecurableObjectProperty<SecurableObjectType>;
  sharedWith: ISecurableObjectProperty<string[]>;
  sharedBy: ISecurableObjectProperty<string[]>;
  key: string;
}

export interface ISecurableObjectProperty<Type> {
  data: Type;
  displayValue: string;
}

export interface IGetExtContentFunc {
    (): Promise<IGetExtContentFuncResponse>;
}

export interface ISecurableObjectStore {
  getExternalContent: IGetExtContentFunc;
}

export interface IExternalSharingDashboardWebPartProps {
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
  noResultsString: string;
  managedPropertyName: string;
}

export interface IExternalSharingDashboardProps {
  store: ISecurableObjectStore;
  contentProps: IExtContentFetcherProps;
}

// export interface ITableProps {
//   items: ISecurableObject[];
// }

export interface ITable {
  columns: ITableRow;
  rows: ITableRow[];
  pageSize: number;
  pageStartIndex: number;
  currentSortOrder: string;
}

export interface ITableRow {
  cells: ITableCell<any>[];
  key: string;
}

export interface ITableCell<Type> {
  sortableData: Type;
  displayData: string;
  key: string;
}
