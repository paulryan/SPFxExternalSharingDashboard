import {
  ControlMode,
  SecurableObjectType
} from "../classes/Enums";

import {
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore
} from "../classes/Interfaces";

export default class MockContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;

  private content: ISecurableObject[] = [
    {
      title: { data: "My first document", displayValue: "My first document" },
      url: { data: "https://www.google.com/01", displayValue: "https://www.google.com/01" },
      type: { data: SecurableObjectType.Document, displayValue: "Document" },
      fileExtension: { data: "docx", displayValue: "docx" },
      lastModifiedTime: { data: new Date(), displayValue: (new Date()).toDateString() },
      sharedWith: { data: ["Paul Ryan"], displayValue: "Paul Ryan" },
      sharedBy: { data: ["Chris O'Brien"], displayValue: "Chris O'Brien" },
      siteID: { data: "1", displayValue: "1" },
      siteTitle: { data: "Team Site", displayValue: "Team Site" },
      crawlTime: { data: new Date(), displayValue: "Never" },
      key: "1"
    },
    {
      title: { data: "My second document", displayValue: "My second document" },
      url: { data: "https://www.google.com/02", displayValue: "https://www.google.com/02" },
      type: { data: SecurableObjectType.Document, displayValue: "Document" },
      fileExtension: { data: "docx", displayValue: "docx" },
      lastModifiedTime: { data: new Date(), displayValue: (new Date()).toDateString() },
      sharedWith: { data: ["Paul Ryan"], displayValue: "Paul Ryan" },
      sharedBy: { data: ["Chris O'Brien"], displayValue: "Chris O'Brien" },
      siteID: { data: "1", displayValue: "1" },
      siteTitle: { data: "Team Site", displayValue: "Team Site" },
      crawlTime: { data: new Date(), displayValue: "Never" },
      key: "2"
    }
  ];

  public constructor (props: IExtContentFetcherProps) {
    this.props = props;
  }

  public getExternalContent(): Promise<IGetExtContentFuncResponse> {
    return new Promise<IGetExtContentFuncResponse>((resolve) => {
        resolve({
          extContent: this.content,
          controlMode: ControlMode.Content,
          message: "Mocked documents",
          mode: this.props.mode,
          scope: this.props.scope
        });
    });
  }
}
