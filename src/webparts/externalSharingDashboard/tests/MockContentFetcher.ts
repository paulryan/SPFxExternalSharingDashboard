import {
  ControlMode,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore,
  SecurableObjectType
} from "../classes/Interfaces";

export default class MockContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;

  private content: ISecurableObject[] = [
    {
      Title: "My first document",
      URL: "https://www.google.com/01",
      Type: SecurableObjectType.Document,
      FileExtension: "docx",
      LastModifiedTime: (new Date()).toDateString(),
      SharedWith: ["Paul Ryan"],
      SharedBy: ["Chris O'Brien"],
      SiteID: "1",
      SiteTitle: "Team Site",
      CrawlTime: "Never",
      key: "1"
    },
    {
      Title: "My second document",
      URL: "https://www.google.com/02",
      Type: SecurableObjectType.Document,
      FileExtension: "pptx",
      LastModifiedTime: (new Date()).toDateString(),
      SharedWith: ["Paul Ryan"],
      SharedBy: ["Chris O'Brien"],
      SiteID: "1",
      SiteTitle: "Team Site",
      CrawlTime: "Never",
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
