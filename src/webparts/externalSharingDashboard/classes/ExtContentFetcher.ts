import {
  ControlMode,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Enums";

import {
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  ISecurableObject,
  ISecurableObjectStore
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

    const rowLimit: number = 500;
    const baseUri: string = self.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    const extContentFql: string = "*"; //"" + self.props.managedProperyName + ":ext"; //":#ext#";

    let modeFql: string = "";
    if (self.props.mode === Mode.AllExtSharedDocuments) {
      modeFql = ""; // Do not need to restrict further
    }
    else if (self.props.mode === Mode.MyExtSharedDocuments) {
      // "MY" should represent things I have created or edited or shared.
      const un: string = self.props.context.pageContext.user.loginName;
      const me: string = un.substring(0, un.indexOf("@")); // TODO: get the query working with @ symbol.. .replace("@", "%40");
      modeFql = ` (ModifiedBy:${me} OR CreatedBy:${me} OR ${self.props.managedProperyName}:${me})`;
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

  private _parseDisplayNameFromExtUserAccountName(extUserAccountName: string): string {
    let extUserDisplayName: string = "";
    if (extUserAccountName) {
      // We want the bit betwee the last index of | and the first index of #
      let startIndex: number = extUserAccountName.lastIndexOf("|");
      if (startIndex < 0) {
        startIndex = 0;
      }
      else {
        startIndex += "|".length;
      }
      let endIndex: number = extUserAccountName.indexOf("#", startIndex);
      if (endIndex < 0) {
        endIndex = extUserAccountName.length;
      }
      extUserDisplayName = extUserAccountName.substring(startIndex, endIndex);
    }
    return extUserDisplayName;
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
              const sharedWithUserDisplayName: string = this._parseDisplayNameFromExtUserAccountName(sharedWithUser);
              sharedWith.push(sharedWithUserDisplayName);
              const sharedByUser: string = sharedWithDetails[sharedWithUser]["LoginName"];
              sharedBy.push(sharedByUser);
            }
          }

          sharedWith.sort();
          sharedBy.sort();

          const lastModifedTime: Date = new Date(doc.LastModifiedTime);
          const crawlTime: Date = new Date(doc.CrawlTime);
          const now: Date = new Date();
          const old: Date = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
          const isCrawlTimeInvalid: boolean = (crawlTime > now || crawlTime < old);

          searchRowsSimplified.push({
            title: { data: doc.Filename, displayValue: doc.Filename},
            fileExtension: { data: doc.FileExtension, displayValue: doc.FileExtension},
            lastModifiedTime: { data: lastModifedTime, displayValue: this.toColloquialDateString(lastModifedTime)},
            siteID: { data: doc.SiteID, displayValue: doc.SiteID},
            siteTitle: { data: doc.SiteTitle, displayValue: doc.SiteTitle},
            url: { data: doc.ServerRedirectedURL || doc.Path, displayValue: doc.ServerRedirectedURL || doc.Path},
            type: { data: SecurableObjectType.Document, displayValue: "Document"},
            sharedBy: { data: sharedBy, displayValue: sharedBy.join(", ")},
            sharedWith: { data: sharedWith, displayValue: sharedWith.join(", ")},
            crawlTime: { data: crawlTime, displayValue: isCrawlTimeInvalid ? "" : this.toColloquialDateString(crawlTime)},
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
  }

  private toShortDateString (date: Date): string {
      // e.g. 18 aug 2015
      //const ds = date.format("ddd, dd MMM yyyy");
      const ds: string = date.toDateString();
      return ds;
  }

  private toColloquialDateString (then: Date): string {
        let returnString: string = this.toShortDateString(then);
        const now: Date = new Date();
        const minsInHr: number = 60;

        const isSameDay: boolean = then.getFullYear() === now.getFullYear() && then.getMonth() === now.getMonth() && then.getDate() === now.getDate();
        if (isSameDay) {
            if (now > then) {
                const totalMinutesAgo: number = (now.getHours() * minsInHr) + now.getMinutes() - (then.getHours() * minsInHr) - then.getMinutes();
                const hoursAgo: number = Math.floor(totalMinutesAgo / minsInHr);
                const minsAgo: number = totalMinutesAgo % minsInHr;

                if (hoursAgo < 1) {
                    returnString = "" + minsAgo + " minutes ago";
                }
                if (hoursAgo === 1) {
                    returnString = "" + hoursAgo + " hour and " + minsAgo + " minutes ago";
                }
                else {
                    returnString = "" + hoursAgo + " hours and " + minsAgo + " minutes ago";
                }
            }
        }
        return returnString;
    };
}
