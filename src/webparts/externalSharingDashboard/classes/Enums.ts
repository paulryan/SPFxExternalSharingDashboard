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
  if (mode === Mode.AllExtSharedDocuments) {
    displayName = "All externally shared documents";
  }
  else if (mode === Mode.MyExtSharedDocuments) {
    displayName = "Documents which I have created, modified, or shared externally";
  }
  return displayName;
}

export enum SPScope {
  Tenant = 1,
  SiteCollection = 2,
  Site = 3
}

export enum Mode {
  AllExtSharedDocuments = 1,
  MyExtSharedDocuments = 2,
  // AllExtSharedContainers = 3,
  // MyExtSharedContainers = 4,
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
  // Library = 2,
  // Web = 3,
  // Site = 4
}

export enum ControlMode {
  Loading = 1,
  Message = 2,
  Content = 3
}
