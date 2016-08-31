import {
  EnvironmentType
} from "@microsoft/sp-client-base";

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneTextField
} from "@microsoft/sp-client-preview";

import * as strings from "mystrings";
import * as React from "react";
import * as ReactDom from "react-dom";

import DocumentDashboard from "./components/DocumentDashboard";

import {
  DisplayType,
  GetDisplayTermForEnumMode,
  GetDisplayTermForEnumSPScope,
  Mode,
  SPScope
} from "./classes/Enums";

import {
  IContentFetcherProps,
  IDocumentDashboardProps,
  IDocumentDashboardWebPartProps,
  ISecurableObjectStore
} from "./classes/Interfaces";

import ContentFetcher from "./classes/ContentFetcher";
import MockContentFetcher from "./tests/MockContentFetcher";

import {
  Logger
} from "./classes/Logger";

export default class DocumentDashboardWebPart extends BaseClientSideWebPart<IDocumentDashboardWebPartProps> {
  private log: Logger;

  public constructor(context: IWebPartContext) {
    super(context);
    this.log = new Logger("DocumentDashboardWebPart");
  }

  public render(): void {
    // Define properties for the Content Fetcher
    const contentFecherProps: IContentFetcherProps = {
      context: this.context,
      scope: this.properties.scope,
      mode: this.properties.mode,
      managedProperyName: this.properties.sharedWithManagedPropertyName,
      crawlTimeManagedPropertyName: this.properties.crawlTimeManagedPropertyName,
      noResultsString: this.properties.noResultsString
    };

    // Create appropriate Content Fectcher class for getting content
    let extContentStore: ISecurableObjectStore;
    if (this.context.environment.type === EnvironmentType.Local || this.context.environment.type === EnvironmentType.Test) {
      extContentStore = new MockContentFetcher(contentFecherProps);
    }
    else {
      extContentStore = new ContentFetcher(contentFecherProps);
    }

    // Create manager ReactElement
    const element: React.ReactElement<IDocumentDashboardProps> = React.createElement(DocumentDashboard, {
      store: extContentStore,
      mode: this.properties.mode,
      scope: this.properties.scope,
      displayType: this.properties.displayType
    });

    // Build the control!
    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Core",
              groupFields: [
                PropertyPaneDropdown("scope", {
                  label: "Where should we look for content?",
                  options: [
                    { key: SPScope.Tenant, text: GetDisplayTermForEnumSPScope(SPScope.Tenant) },
                    { key: SPScope.SiteCollection, text: GetDisplayTermForEnumSPScope(SPScope.SiteCollection) },
                    { key: SPScope.Site, text: GetDisplayTermForEnumSPScope(SPScope.Site) }
                  ]
                }),
                PropertyPaneDropdown("mode", {
                  label: "What type content do you want to see?",
                  options: [
                    { key: Mode.AllExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.AllExtSharedDocuments) },
                    { key: Mode.MyExtSharedDocuments, text: GetDisplayTermForEnumMode(Mode.MyExtSharedDocuments) },
                    { key: Mode.AllDocuments, text: GetDisplayTermForEnumMode(Mode.AllDocuments) },
                    { key: Mode.MyDocuments, text: GetDisplayTermForEnumMode(Mode.MyDocuments) }
                    // { key: Mode.AllExtSharedContainers, text: "All externally shared sites, libraries, and folders" },
                    // { key: Mode.MyExtSharedContainers, text: "Sites, libraries, and folders which I have shared externally" }
                  ]
                }),
                PropertyPaneDropdown("displayType", {
                  label: "How do you want the results rendered?",
                  options: [
                    { key: DisplayType.Table, text: "As a table" },
                   // { key: DisplayType.Tree, text: "Hierarchically" },
                    { key: DisplayType.BySite, text: "Charted by site" },
                    { key: DisplayType.ByUser, text: "Charted by user" },
                    { key: DisplayType.OverTime, text: "Charted over time" }
                  ]
                })
              ]
            },
            {
              groupName: "Other",
              groupFields: [
                PropertyPaneTextField("noResultsString", {
                  label: "What message should we display when there are no results?"
                }),
                PropertyPaneTextField("sharedWithManagedPropertyName", {
                  label: "What is the name of the queryable Managed Property containing shared with details?",
                  description: `This property must be configured as such:
                                Text, Multi, Queryable, Retrievable, and be mapped to 'ows_SharedWithDetails'`
                }),
                PropertyPaneTextField("crawlTimeManagedPropertyName", {
                  label: "What is the name of the Managed Property containing crawl time details?",
                  description: `This property must be configured as such:
                                Text, Retrievable, and be mapped to 'Internal:323'`
                }),
                PropertyPaneLabel("labelproperty01", {
                  text: "Use the following link to download a search schema file to import the above managed properties:"
                }),
                PropertyPaneLink("linkproperty01", {
                  href: "https://www.bing.com",
                  text: "Search schema"
                }),
                PropertyPaneLabel("labelproperty02", {
                  text: "_"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
