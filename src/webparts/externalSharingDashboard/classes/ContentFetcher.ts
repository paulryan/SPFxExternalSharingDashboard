import {
  ControlMode,
  Mode,
  SPScope,
  SecurableObjectType
} from "./Enums";

import {
  IContentFetcherProps,
  IGetContentFuncResponse,
  ISearchResponse,
  ISecurableObject,
  ISecurableObjectStore
} from "./Interfaces";

import {
  EnsureBracesOnGuidString,
  ParseDisplayNameFromExtUserAccountName,
  ToColloquialDateString,
  TransformSearchResponse
} from "./Utilities";

import {
  Logger
} from "./Logger";

export default class ContentFetcher implements ISecurableObjectStore {

  public props: IContentFetcherProps;
  private log: Logger;

  public constructor(props: IContentFetcherProps) {
    this.props = props;
    this.log = new Logger("ContentFetcher");
  }

  public getContent(): Promise<IGetContentFuncResponse> {
    const self: ContentFetcher = this;
    self.log.logInfo("getContent()");

    const rowLimit: number = 500;
    const baseUri: string = self.props.context.pageContext.web.absoluteUrl + "/_api/search/query";

    // "MY" should represent things I have created or edited or shared.
    const un: string = self.props.context.pageContext.user.loginName;
    const me: string = un.substring(0, un.indexOf("@")); // TODO: get the query working with @ symbol.. .replace("@", "%40");
    let myFql: string = `(ModifiedBy:${me} OR CreatedBy:${me}`;
    if (self.props.managedProperyName) {
      myFql = `(ModifiedBy:${me} OR CreatedBy:${me} OR ${self.props.managedProperyName}:${me})`;
    }

    let modeFql: string = "";
    if (self.props.mode === Mode.AllExtSharedDocuments) {
      modeFql = "" + self.props.managedProperyName + ":ext";
    }
    else if (self.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = `${self.props.managedProperyName}:ext ${myFql}`;
    }
    else if (self.props.mode === Mode.AllDocuments) {
      modeFql = "ContentClass:STS_ListItem";
    }
    else if (self.props.mode === Mode.MyDocuments) {
      // TODO: Can we get Graphy in a meaningful way?
      modeFql = `ContentClass:STS_ListItem ${myFql}`;
    }
    else {
      self.log.logError("Unsupported mode: " + self.props.mode);
      return null;
    }

    let scopeFql: string = "";
    if (self.props.scope === SPScope.SiteCollection) {
      scopeFql = "SiteId:" + EnsureBracesOnGuidString(self.props.context.pageContext.site.id.toString());
    }
    else if (self.props.scope === SPScope.Site) {
      scopeFql = "WebId:" + EnsureBracesOnGuidString(self.props.context.pageContext.web.id.toString());
    }
    else if (self.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      self.log.logError("Unsupported scope: " + self.props.scope);
      return null;
    }

    const selectPropsArray: string[] = ["LastModifiedTime", "Title", "Filename", "ServerRedirectedURL", "Path", "FileExtension", "UniqueID", "SharedWithDetails", "SiteTitle", "SiteID"];
    if (this.props.crawlTimeManagedPropertyName) {
      selectPropsArray.push(this.props.crawlTimeManagedPropertyName);
    }

    const queryText: string = "querytext='" + scopeFql + " " +  modeFql + "'";
    const selectProps: string = "selectproperties='" + selectPropsArray.join(",") + "'";
    const finalUri: string = baseUri + "?" + queryText + "&" + selectProps;

    const prom: Promise<IGetContentFuncResponse> = new Promise<IGetContentFuncResponse>((resolve: any, reject: any) => {
      self.queryForAllItems(finalUri, 0, rowLimit, null, resolve, reject);
    });
    return prom;
  }

  private queryForAllItems(uri: string, startIndex: number, rowlimit: number, response: IGetContentFuncResponse, resolve: any, reject: any): void {

    const pagedUri: string = uri + "&startRow=" + startIndex + "&rowLimit=" + rowlimit;
    this.log.logInfo("Submitting request to " + pagedUri);

    const headers: Headers = new Headers();
    headers.append("odata-version", "3.0");
    this.props.context.httpClient.get(pagedUri, { headers: headers })
      .then((r1: Response) => {
        this.log.logInfo("Recieved response from " + pagedUri);
        if (r1.ok) {
          r1.json().then((r) => {
            const searchResponse: ISearchResponse = TransformSearchResponse(r);
            const currentResponse: IGetContentFuncResponse = this.transformSearchResultsToResponseObject(searchResponse, this.props.noResultsString);
            if (response) {
              response.results.push(...currentResponse.results);
            }
            else {
              response = currentResponse;
            }
            let getAnotherPage: boolean = false;
            if (!currentResponse.isError) {
              // Get the next page if results === rowlimit
              const rowCount: number = r.PrimaryQueryResult.RelevantResults.RowCount;
              if (rowCount === rowlimit && startIndex + rowCount < r.PrimaryQueryResult.RelevantResults.TotalRows) {
                getAnotherPage = true;
              }
            }
            if (getAnotherPage) {
              // Recursive call
              this.log.logInfo("Fetching an additional page of results");
              this.queryForAllItems(uri, startIndex + rowlimit, rowlimit, response, resolve, reject);
            }
            else {
              resolve(response);
            }
          });
        }
        else {
          reject({
            extContent: [],
            controlMode: ControlMode.Message,
            message: r1.statusText
          });
        }
      }, (error: any) => {
        reject({
          extContent: [],
          controlMode: ControlMode.Message,
          message: "Sorry, there was an error submitting the request"
        });
      });
  }

  private transformSearchResultsToResponseObject(searchResponse: ISearchResponse, noResultsString: string): IGetContentFuncResponse {
    let isError: boolean = false;
    let message: string = "";
    const securableObjects: ISecurableObject[] = [];
    if (searchResponse.isSuccess) {
      if (searchResponse && searchResponse.isSuccess && searchResponse.results.length > 0) {
        try {
          searchResponse.results.forEach((doc: any) => {
            // Parse out SharedWithDetails
            const sharedWith: string[] = [];
            const sharedBy: string[] = [];
            if (doc.SharedWithDetails) {
              const sharedWithDetails: any = JSON.parse(doc.SharedWithDetails);
              for (const sharedWithUser in sharedWithDetails) {
                if (sharedWithDetails.hasOwnProperty(sharedWithUser)) {
                  const sharedWithUserDisplayName: string = ParseDisplayNameFromExtUserAccountName(sharedWithUser);
                  sharedWith.push(sharedWithUserDisplayName);
                  const sharedByUser: string = sharedWithDetails[sharedWithUser]["LoginName"];
                  sharedBy.push(sharedByUser);
                }
              }
              sharedWith.sort();
              sharedBy.sort();
            }

            // Parse CrawlTime if managed property provided and populated
            let crawlTime: Date = null;
            let isCrawlTimeInvalid: boolean = true;
            if (this.props.crawlTimeManagedPropertyName) {
              const crawlTimeString: string = doc[this.props.crawlTimeManagedPropertyName];
              if (crawlTimeString) {
                crawlTime = new Date(crawlTimeString);
                const now: Date = new Date();
                const old: Date = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
                isCrawlTimeInvalid = (crawlTime > now || crawlTime < old);
              }
            }

            // Parse last modified date
            let lastModifedTime: Date = null;
            let isLastModifiedTimeInvalid: boolean = true;
            if (doc.LastModifiedTime) {
              lastModifedTime = new Date(doc.LastModifiedTime);
              isLastModifiedTimeInvalid = false;
            }

            // Create ISecurableObject from search results
            securableObjects.push({
              title: { data: doc.Filename, displayValue: doc.Filename },
              fileExtension: { data: doc.FileExtension, displayValue: doc.FileExtension },
              lastModifiedTime: { data: lastModifedTime, displayValue: isLastModifiedTimeInvalid ? "" : ToColloquialDateString(lastModifedTime) },
              siteID: { data: doc.SiteID, displayValue: doc.SiteID },
              siteTitle: { data: doc.SiteTitle, displayValue: doc.SiteTitle },
              url: { data: doc.ServerRedirectedURL || doc.Path, displayValue: doc.ServerRedirectedURL || doc.Path },
              type: { data: SecurableObjectType.Document, displayValue: "Document" },
              sharedBy: { data: sharedBy, displayValue: sharedBy.join(", ") },
              sharedWith: { data: sharedWith, displayValue: sharedWith.join(", ") },
              crawlTime: { data: crawlTime, displayValue: isCrawlTimeInvalid ? "" : ToColloquialDateString(crawlTime) },
              key: doc.UniqueID || doc.Path
            });
          });
        }
        catch (e) {
          isError = true;
          message = "Sorry, there was an error parsing the response.";
          this.log.logError(message, e.message);
        }
      }

      if (!isError && securableObjects.length < 1) {
        message = noResultsString;
      }
    }
    else {
      isError = true;
      message = searchResponse.message;
    }

    return {
      results: securableObjects,
      isError: isError,
      message: message
    };
  }
}
