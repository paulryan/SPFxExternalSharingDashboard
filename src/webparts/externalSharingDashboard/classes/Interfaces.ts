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

export interface IContentFetcherProps {
  context: IWebPartContext;
  scope: SPScope;
  mode: Mode;
  managedProperyName: string;
  crawlTimeManagedPropertyName: string;
  noResultsString: string;
}

export interface IGetContentFuncResponse {
  results: ISecurableObject[];
  message: string;
  isError: boolean;
}

export interface IDocumentDashboardProps {
  store: ISecurableObjectStore;
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
}

export interface IDocumentDashboardState {
  results: ISecurableObject[];
  message: string;
  controlMode: ControlMode;
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
}

export interface ISecurableObject {
  title: ISecurableObjectProperty<string>;
  fileExtension: ISecurableObjectProperty<string>;
  lastModifiedTime: ISecurableObjectProperty<Date>;
  siteTitle: ISecurableObjectProperty<string>;
  siteID: ISecurableObjectProperty<string>;
  crawlTime: ISecurableObjectProperty<Date>;
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

export interface IGetContentFunc {
    (): Promise<IGetContentFuncResponse>;
}

export interface ISecurableObjectStore {
  getContent: IGetContentFunc;
}

export interface IDocumentDashboardWebPartProps {
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
  noResultsString: string;
  sharedWithManagedPropertyName: string;
  crawlTimeManagedPropertyName: string;
}

export interface ISearchResponse {
  results: any[];
  rowCount: number;
  totalRows: number;
  totalRowsIncludingDuplicates: number;
  isSuccess: boolean;
  message: string;
}

export interface IChart {
  items: IChartItem[];
  columnIndexToGroupUpon: number;
}

export interface IChartItem {
  label: string;
  value: number;
}

export interface ITable {
  columns: ITableRow;
  rows: ITableRow[];
  pageSize: number;
  pageStartIndex: number;
  currentSort: number;
  currentSortDescending: boolean;
}

export interface ITableRow {
  cells: ITableCell<any>[];
  key: string;
}

export interface ITableCell<Type> {
  sortableData: Type;
  displayData: string;
  href: string;
  onClick?: any;
  key: string;
}
