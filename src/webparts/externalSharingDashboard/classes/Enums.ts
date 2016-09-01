export function GetDisplayTermForEnumSPScope(scope: SPScope): string {
  let displayName: string = "";
  if (scope === SPScope.Tenant) {
    displayName = "Anywhere in the entire tenancy";
  }
  else if (scope === SPScope.SiteCollection) {
    displayName = "Within this site collection";
  }
  else if (scope === SPScope.Site) {
    displayName = "Within this site (or child sites)";
  }
  return displayName;
}

export function GetDisplayTermForEnumMode(mode: Mode): string {
  let displayName: string = "";
  if (mode === Mode.AllDocuments) {
    displayName = "All documents";
  }
  else if (mode === Mode.MyDocuments) {
    displayName = "My documents"; //"Documents which I have created or modfied";
  }
  else if (mode === Mode.AllExtSharedDocuments) {
    displayName = "All externally shared documents";
  }
  else if (mode === Mode.MyExtSharedDocuments) {
    displayName = "My externally shared documents"; //"Externally shared documents which I have created, modified, or shared";
  }
  else if (mode === Mode.AllAnonSharedDocuments) {
    displayName = "All anonymously shared documents";
  }
  else if (mode === Mode.MyAnonSharedDocuments) {
    displayName = "My anonymously shared documents"; //"Anonymously shared documents which I have created, modified, or shared";
  }
  return displayName;
}

export enum SPScope {
  Tenant = 1,
  SiteCollection = 2,
  Site = 3
}

export enum Mode {
  AllDocuments = 1,
  MyDocuments = 2,
  AllExtSharedDocuments = 3,
  MyExtSharedDocuments = 4,
  AllAnonSharedDocuments = 5,
  MyAnonSharedDocuments = 6
}

export enum DisplayType {
  Table = 1,
  BySite = 2,
  ByUser = 3,
  OverTime = 4
}

export enum SecurableObjectType {
  Document = 1,
  // Library = 2,
  // Web = 3,
  // Site = 4
}

export enum ControlMode {
  Loading = 1,
  Message = 2,
  Content = 3
}
