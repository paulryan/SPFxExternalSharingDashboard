import {
  ISecurableObjectStore,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  SecurableObjectType,
  ControlMode
} from "../classes/Interfaces";

export default class MockContentFetcher implements ISecurableObjectStore {

  public timeStamp: number;

  public constructor (props: IExtContentFetcherProps) {
    this.timeStamp = -1;
  }

  public getAllExtDocuments(): Promise<IGetExtContentFuncResponse> {
    return new Promise<IGetExtContentFuncResponse>((resolve) => {
        this.timeStamp = (new Date()).getTime();
        resolve({
          extContent: this._content,
          controlMode: ControlMode.Content,
          message: "Mocked documents",
          timeStamp: this.timeStamp
        });
    });
  }

  private _content: ISecurableObject[] = [
    {
      Title: "My first document",
      URL: "https://www.google.com/01",
      Type: SecurableObjectType.Document,
      FileExtension: "docx",
      LastModifiedTime: (new Date()).toDateString(),
      SharedWith: "Paul Ryan",
      SharedBy: "Chris O'Brien",
      key: "1"
    },
    {
      Title: "My second document",
      URL: "https://www.google.com/02",
      Type: SecurableObjectType.Document,
      FileExtension: "pptx",
      LastModifiedTime: (new Date()).toDateString(),
      SharedWith: "Paul Ryan",
      SharedBy: "Chris O'Brien",
      key: "2"
    }
  ];
}
