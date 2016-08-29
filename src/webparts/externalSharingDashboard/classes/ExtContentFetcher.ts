import {
  ControlMode,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Interfaces";

import {
  Logger
} from "./Logger";

export default class ExtContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;
  private log: Logger;

  public constructor (props: IExtContentFetcherProps) {
    this.props = props;
    this.log = new Logger("ExtContentFetcher");
  }

  public getExternalContent(): Promise<IGetExtContentFuncResponse> {
    const self: ExtContentFetcher = this;
    self.log.logInfo("getExternalContent()");

    // TODO : do some clever caching
    const rowLimit: number = 500; // TODO : we need to get many pages with this etc..
    const baseUri: string = self.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    const extContentFql: string = "" + self.props.managedProperyName + ":ext"; //":#ext#";

    let modeFql: string = "";
    if (self.props.mode === Mode.AllExtSharedDocuments) {
      modeFql = ""; // Do not need to restrict further
    }
    else if (self.props.mode === Mode.MyExtSharedDocuments) {
      // "MY" should represent things I have created or edited or shared.
      const un: string = self.props.context.pageContext.user.loginName;
      const me: string = un.substring(0, un.indexOf("@")); // TODO: get the query working with @ symbol.. .replace("@", "%40");
      modeFql = ` (ModifiedBy:${me} OR CreatedBy:${me} OR ${self.props.managedProperyName}:${me})`; // Do not need to restrict further
      //modeFql = ` ${self.props.managedProperyName}:${me}`;
    }
    else {
      self.log.logError("Unsupported mode: " + self.props.mode);
      return null;
    }

    let scopeFql: string = "";
    if (self.props.scope === SPScope.SiteCollection) {
      scopeFql = " SiteId:" + self.props.context.pageContext.site.id.toString();
    }
    else if (self.props.scope === SPScope.Site) {
      scopeFql = " WebId:{" + self.props.context.pageContext.web.id.toString() + "}";
    }
    else if (self.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      self.log.logError("Unsupported scope: " + self.props.scope);
      return null;
    }

    const queryText: string = "querytext='" + extContentFql + scopeFql + modeFql + "'";
    const selectProps: string = "selectproperties='Title,Filename,ServerRedirectedURL,Path,FileExtension,UniqueID,SharedWithDetails,SiteTitle,SiteID,CrawlTime'";
    const finalUri: string = baseUri + "?" + queryText + "&" + selectProps;

    const prom: Promise<IGetExtContentFuncResponse> = new Promise<IGetExtContentFuncResponse>((resolve: any, reject: any) => {
      self._queryForAllItems(finalUri, 0, rowLimit, null, resolve, reject);
    });
    return prom;
  }

  private _queryForAllItems(uri: string, startIndex: number, rowlimit: number, results: IGetExtContentFuncResponse,
                            resolve: any, reject: any): void {

    const pagedUri: string = uri + "&startRow=" + startIndex + "&rowLimit=" + rowlimit;
    this.log.logInfo("Submitting request to " + pagedUri);

    this.props.context.httpClient.get(pagedUri)
      .then(
        (r1: Response) => {
        r1.json().then((r) => {
          this.log.logInfo("Recieved response from " + pagedUri);
          const errorMsg: string = r.error ? r.error.message : r.message;
          if (errorMsg) {
            reject({
              extContent: [],
              controlMode: ControlMode.Message,
              message: errorMsg
            });
          }
          else {
            const finalResults: IGetExtContentFuncResponse = this._transformSearchResults(r, this.props.noResultsString);
            if (results) {
              results.extContent.push(...finalResults.extContent);
            }
            else {
              results = finalResults;
            }
            let getAnotherPage: boolean = false;
            if (finalResults.controlMode === ControlMode.Content) {
              // Get the next page if results === rowlimit
              const rowCount: number = r.PrimaryQueryResult.RelevantResults.RowCount;
              if (rowCount === rowlimit && startIndex + rowCount < r.PrimaryQueryResult.RelevantResults.TotalRows) {
                getAnotherPage = true;
              }
            }

            if (getAnotherPage) {
              // Recursive call
              this._queryForAllItems(uri, startIndex + rowlimit, rowlimit, results, resolve, reject);
            }
            else {
              resolve(results);
            }
          }
        });
      }, (error: any) => {
        reject({
          extContent: [],
          controlMode: ControlMode.Message,
          message: "Sorry, there was an error submitting the request"
        });
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

          let sharedWithDetails: any = {};
          if (doc.SharedWithDetails) {
            sharedWithDetails = JSON.parse(doc.SharedWithDetails);
          }
          const sharedWith: string[] = [];
          const sharedBy: string[] = [];

          for (const sharedWithUser in sharedWithDetails) {
            if (sharedWithDetails.hasOwnProperty(sharedWithUser)) {
              sharedWith.push(sharedWithUser);
              const sharedByUser: string = sharedWithDetails[sharedWithUser]["LoginName"];
              sharedBy.push(sharedByUser);
            }
          }

          searchRowsSimplified.push({
            Title: doc.Filename,
            FileExtension: doc.FileExtension,
            LastModifiedTime: "", // doc.LastModifiedTime,
            SiteID: doc.SiteID,
            SiteTitle: doc.SiteTitle,
            URL: doc.ServerRedirectedURL || doc.Path,
            Type: SecurableObjectType.Document,
            SharedBy: sharedBy,
            SharedWith: sharedWith,
            CrawlTime: doc.CrawlTime,
            key: doc.UniqueID || doc.Path
          });
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
      mode: this.props.mode,
      scope: this.props.scope
    };
  };
}
