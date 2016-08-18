import * as React from "react";

import {
  IExternalSharingDashboardProps,
  IGetExtContentFuncResponse,
  ControlMode,
  ISecurableObject
} from "../classes/Interfaces";

import {
  Logger
} from "../classes/Logger";

import {
  Spinner,
  SpinnerType,
  Label
} from "office-ui-fabric-react";

import Table from "./Table";

interface ITableProps {
  items: ISecurableObject[];
}

interface ITableRowProps {
  item: ISecurableObject;
}

export default class ExternalSharingDashboard extends React.Component<IExternalSharingDashboardProps, IGetExtContentFuncResponse> {
  private log: Logger;

  constructor() {
    super();
    this.log = new Logger("ExternalSharingDashboard");
  }

  public componentWillMount(): void {
    this.log.logInfo("componentWillMount");
    this._setStateToLoading();
  }

  public componentDidMount(): void {
    this.log.logInfo("componentDidMount");
    this._updateState();
  }

  private updatedOnce: boolean = false;

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    if (this.state.timeStamp === this.props.store.timeStamp) {
      // Do nothing, as the data will be the same
    } else {
      if (!this.updatedOnce) {
        this.updatedOnce = true;
        this._updateState();
      }
    }
  }

  private _updateState(): void {
    this.log.logInfo("_updateState");
    this._setStateToLoading();
    this.props.store.getAllExtDocuments()
    .then((r) => {
      this.setState(r);
      this.log.logInfo("_setStateToContent");
    });
  }

  private _setStateToLoading(): void {
    this.log.logInfo("_setStateToLoading");
    this.setState({
      extContent: [],
      controlMode: ControlMode.Loading,
      message: "Working on it...",
      timeStamp: (new Date()).getTime()
    });
  }

  public render(): JSX.Element {
    this.log.logInfo("render");
    if (this.state && this.state.controlMode === ControlMode.Loading) {
      return (
        <Spinner type={ SpinnerType.large } label={this.state.message} />
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      return (
        <Label>{this.state.message}</Label>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {
      return (
        <Table items={this.state.extContent} />
      );
    }
    else {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
    }
  }
}
