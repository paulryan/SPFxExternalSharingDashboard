import * as React from "react";

import {
  ControlMode,
  IExternalSharingDashboardProps,
  IGetExtContentFuncResponse,
  ISecurableObject
} from "../classes/Interfaces";

import {
  Logger
} from "../classes/Logger";

import {
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
  private updateStateInProgress: boolean = false;

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

  public componentWillReceiveProps(): void {
    this.log.logInfo("componentWillReceiveProps");
    this._setStateToLoading();
  }

  public shouldComponentUpdate(nextProps: IExternalSharingDashboardProps, nextState: IGetExtContentFuncResponse): boolean {
    return !nextState || nextProps.contentProps.mode !== this.state.mode
      || nextProps.contentProps.scope !== this.state.scope;
  }

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    this._updateState();
  }

  public render(): JSX.Element {
    this.log.logInfo("render");
    if (this.state && this.state.controlMode === ControlMode.Loading) {
      return (
        //<Spinner type={ SpinnerType.large } label={this.state.message} />
        <div className="ms-font-l">{this.state.message}</div>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      return (
        <Label>{this.state.message}</Label>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {
      return (
        <div>
        <div className="ms-font-xxl">{this.state.mode}</div>
        <div className="ms-font-l">{this.state.scope}</div>
        <Table items={this.state.extContent} />
        </div>
      );
    }
    else if (this.state && this.state.controlMode) {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
      return (<div className="ms-font-l">Error!</div>);
    }
    else {
      this.log.logError(`State is undefined`);
      //this.log.logInfo(`State is undefined`);
      // This will occur is the state is not set in componentWillMount
      return (
        <div className="ms-font-l">Error!</div>
      );
    }
  }

  private _updateState(): void {
      if (!this.updateStateInProgress) {
        this.updateStateInProgress = true;
        this.props.store.getExternalContent()
        .then((r) => {
          this.setState(r);
          this.updateStateInProgress = false;
        });
      }
  }

  private _setStateToLoading(): void {
    this.setState({
      extContent: [],
      controlMode: ControlMode.Loading,
      message: "Working on it...",
      mode: -1,
      scope: -1
    });
  }
}
