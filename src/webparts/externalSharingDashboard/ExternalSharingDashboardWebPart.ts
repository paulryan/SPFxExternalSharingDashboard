import {
  EnvironmentType
} from "@microsoft/sp-client-base";

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from "@microsoft/sp-client-preview";

import * as React from "react";
import * as ReactDom from "react-dom";

import * as strings from "mystrings";

import ExternalSharingDashboard from "./components/ExternalSharingDashboard";

import {
  DisplayType,
  IExtContentFetcherProps,
  IExternalSharingDashboardProps,
  IExternalSharingDashboardWebPartProps,
  ISecurableObjectStore,
  Mode,
  SPScope
} from "./classes/Interfaces";

import ExtContentFetcher from "./classes/ExtContentFetcher";
import MockContentFetcher from "./tests/MockContentFetcher";

import {
  Logger
} from "./classes/Logger";

export default class ExternalSharingDashboardWebPart extends BaseClientSideWebPart<IExternalSharingDashboardWebPartProps> {
  private log: Logger;

  public constructor(context: IWebPartContext) {
    super(context);
    this.log = new Logger("ExternalSharingDashboardWebPart");
  }

  public render(): void {
    // Define properties for the Content Fetcher
    const contentFecherProps: IExtContentFetcherProps = {
      context: this.context,
      scope: this.properties.scope,
      mode: this.properties.mode,
      managedProperyName: this.properties.managedPropertyName,
      noResultsString: this.properties.noResultsString
    };

    // Create appropriate Content Fectcher class for getting content
    let extContentStore: ISecurableObjectStore;
    if (this.context.environment.type === EnvironmentType.Local || this.context.environment.type === EnvironmentType.Test) {
      extContentStore = new MockContentFetcher(contentFecherProps);
    }
    else {
      extContentStore = new ExtContentFetcher(contentFecherProps);
    }

    // Create appropriate ReactElement for displaying content
    let element: React.ReactElement<IExternalSharingDashboardProps> = null;
    if (this.properties.displayType === DisplayType.Table) {
      element = React.createElement(ExternalSharingDashboard, { store: extContentStore });
    }
    else {
      this.log.logError("Unsupported display type: " + this.properties.displayType);
      return null;
    }

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
                  label: "Where should we look for externally shared content?",
                  options: [
                    { key: SPScope.Tenant, text: "Across the entire tenancy" },
                    { key: SPScope.SiteCollection, text: "Only within this site collection" },
                    { key: SPScope.Site, text: "Only in this site (and in child sites)" }
                  ]
                }),
                PropertyPaneDropdown("mode", {
                  label: "What type content do you want to see?",
                  options: [
                    { key: Mode.AllExtSharedDocuments, text: "All externally shared documents" },
                    { key: Mode.MyExtSharedDocuments, text: "Documents which I have shared externally" },
                    { key: Mode.AllExtSharedContainers, text: "All externally shared sites, libraries, and folders" },
                    { key: Mode.MyExtSharedContainers, text: "Sites, libraries, and folders which I have shared externally" }
                  ]
                }),
                PropertyPaneDropdown("displayType", {
                  label: "How do you want the results rendered?",
                  options: [
                    { key: DisplayType.Table, text: "As a table" },
                    { key: DisplayType.Tree, text: "Hierarchically" },
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
                PropertyPaneTextField("managedPropertyName", {
                  label: "What is the name of the Managed Property with shared details?",
                  description: `This property must be configured as such:
                                Text, Multi, Queryable, Retrievable, and be mapped to 'ows_SharedWithDetails'`
                }),
                PropertyPaneTextField("noResultsString", {
                  label: "What message should we display when there are no results?"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
