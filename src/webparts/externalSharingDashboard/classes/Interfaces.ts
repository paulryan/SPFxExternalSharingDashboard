import {
  IWebPartContext
} from "@microsoft/sp-client-preview";

export enum SPScope {
  Tenant = 1,
  SiteCollection = 2,
  Site = 3
}

export enum Mode {
  AllExtSharedDocuments = 1,
  MyExtSharedDocuments = 2,
  AllExtSharedContainers = 3, // May not be possible
  MyExtSharedContainers = 4, // May not be possible
}

export enum DisplayType {
  Table = 1,
  Tree = 2,
  BySite = 3,
  ByUser = 4,
  OverTime = 5
}

export enum SecurableObjectType {
  Document = 1,
  Library = 2,
  Web = 3,
  Site = 4
}

export enum ControlMode {
  Loading = 1,
  Message = 2,
  Content = 3
}

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
  timeStamp: number;
}

export interface ISecurableObject {
  Title: string;
  URL: string;
  Type: SecurableObjectType;
  FileExtension: string;
  LastModifiedTime: string;
  SharedWith: string;
  SharedBy: string;
  key: string;
}

export interface IGetExtContentFunc {
    (): Promise<IGetExtContentFuncResponse>;
}

export interface ISecurableObjectStore {
  timeStamp: number;
  getAllExtDocuments: IGetExtContentFunc;
  // getMyExtDocuments: IGetExtContentFunc;
  // getAllExtNonDocuments: IGetExtContentFunc;
  // getMyExtNonDocuments: IGetExtContentFunc;
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
}

export interface ITableProps {
  items: ISecurableObject[];
}
