import {
  ControlMode,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore,
  Mode,
  SPScope
} from "./Interfaces";

import {
  Logger
} from "./Logger";

export default class ExtContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;
  public timeStamp: number;
  private log: Logger;

  public constructor (props: IExtContentFetcherProps) {
    this.props = props;
    this.log = new Logger("ExtContentFetcher");
    this.timeStamp = -1;
  }

  public getAllExtDocuments(): Promise<IGetExtContentFuncResponse> {
    this.log.logInfo("getAllExtDocuments()");

    // TODO : do some clever caching
    const rowLimit: number = 500; // TODO : we need to get many pages with this etc..
    const baseUri: string = this.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    const extContentFql: string = "" + this.props.managedProperyName + ":ext"; //":#ext#";

    let scopeFql: string = "";
    if (this.props.scope === SPScope.SiteCollection) {
      scopeFql = " SiteId:" + this.props.context.pageContext.site.id.toString();
    }
    else if (this.props.scope === SPScope.Site) {
      scopeFql = " WebId:" + this.props.context.pageContext.web.id.toString();
    }
    else if (this.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      this.log.logError("Unsupported scope: " + this.props.scope);
      return null;
    }

    // "MY" should represent things I have created or edited or shared.
    let modeFql: string = "";
    if (this.props.mode === Mode.AllExtSharedDocuments || this.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = ""; //" IsDocument=1"; // TODO: Do something better than this
    }
    else if (this.props.mode === Mode.AllExtSharedContainers || this.props.mode === Mode.MyExtSharedContainers) {
      modeFql = " IsDocument=0"; // TODO: Do something better than this
    }
    else {
      this.log.logError("Unsupported mode: " + this.props.mode);
      return null;
    }

    const queryText: string = "querytext='" + extContentFql + scopeFql + modeFql + "'";
    const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension'";
    const finalUri: string = baseUri + "?" + queryText + "&" + selectProps + "&rowlimit=" + rowLimit;

    this.log.logInfo("Submitting request to " + finalUri);

    // TODO: Tidy this up
    return this.props.context.httpClient.get(finalUri)
      .then(
        (r1: Response) => {
        return r1.json().then((r) => {
          this.log.logInfo("Recieved response from " + finalUri);
          const errorMsg: string = r.error ? r.error.message : r.message;
          if (errorMsg) {
            return {
              extContent: [],
              controlMode: ControlMode.Message,
              message: errorMsg,
              timeStamp: (new Date()).getTime()
            };
          }
          else {
            const finalResults: IGetExtContentFuncResponse = this._transformSearchResults(r, this.props.noResultsString);
            this.timeStamp = finalResults.timeStamp;
            return finalResults;
          }
        });
      }, (error: any) => {
        return {
          extContent: [],
          controlMode: ControlMode.Message,
          message: "Sorry, there was an error submitting the request",
          timeStamp: (new Date()).getTime()
        };
      });
  }

  private _transformSearchResults(response: any, noResultsString: string): IGetExtContentFuncResponse {
    // Simplify the data strucutre
    let shouldShowMessage: boolean = false;
    let message: string = "";

    const searchRowsSimplified: ISecurableObject[] = [];

    if (response.PrimaryQueryResult) {
      try {
        const searchRows: any[] = response.PrimaryQueryResult.RelevantResults.Table.Rows;
        searchRows.forEach((d: any) => {
          const doc: any = {};
          d.Cells.forEach((c: any) => {
            doc[c.Key] = c.Value;
          });
          // TODO : convert to ISecurableObject
          searchRowsSimplified.push(doc);
        });
      } catch (e) {
        // TODO: log something?
        shouldShowMessage = true;
        message = "Sorry, there was an error parsing the response";
      }
    }

    if (searchRowsSimplified.length < 1) {
      shouldShowMessage = true;
      message = noResultsString;
    }

    return {
      extContent: searchRowsSimplified,
      controlMode: shouldShowMessage ? ControlMode.Message : ControlMode.Content,
      message: message,
      timeStamp: (new Date()).getTime()
    };
  };
}
