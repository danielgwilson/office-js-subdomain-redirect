import React, { Component } from "react";
import "./App.css";
import { CompoundButton } from "office-ui-fabric-react/lib/Button";

class App extends Component<any, any> {
  constructor(props: any) {
    super(props);
    this.state = {
      range: "",
      hostname: window.location.hostname
    };
    this._redirect = this._redirect.bind(this);
    this._getRange = this._getRange.bind(this);
  }

  render() {
    let headerClassName;
    let nav;
    let actions;
    if (this.state.hostname == "puddle.dgw.cloud") {
      headerClassName = "App-header-puddle";
      nav = (
        <div>
          <div>
            <CompoundButton
              className="Compound-Button"
              primary={true}
              text="Contoso"
              secondaryText="Log into Contoso's analytics cloud."
              onClick={() => this._redirect("https://contoso.dgw.cloud")}
            />
          </div>
          <div>
            <CompoundButton
              className="Compound-Button"
              primary={true}
              text="ACME"
              secondaryText="Log into ACME's analytics cloud."
              onClick={() => this._redirect("https://acme.dgw.cloud")}
            />
          </div>
        </div>
      );
      actions = undefined;
    } else if (this.state.hostname == "contoso.dgw.cloud") {
      headerClassName = "App-header-contoso";
      nav = (
        <div>
          <CompoundButton
            className="Compound-Button"
            text="Log out"
            secondaryText="Log out of the analytics cloud."
            onClick={() => this._redirect("https://puddle.dgw.cloud")}
          />
        </div>
      );
      actions = (
        <div>
          <CompoundButton
            className="Compound-Button"
            primary={true}
            text="Get range"
            secondaryText="Get the range of the current selection."
            onClick={this._getRange}
          />
          <p>{this.state.range}</p>
        </div>
      );
    } else if (this.state.hostname == "acme.dgw.cloud") {
      headerClassName = "App-header-acme";
      nav = (
        <div>
          <CompoundButton
            className="Compound-Button"
            text="Log out"
            secondaryText="Log out of the analytics cloud."
            onClick={() => this._redirect("https://puddle.dgw.cloud")}
          />
        </div>
      );
      actions = (
        <div>
          <CompoundButton
            className="Compound-Button"
            primary={true}
            text="Get range"
            secondaryText="Get the range of the current selection."
            onClick={this._getRange}
          />
          <p>{this.state.range}</p>
        </div>
      );
    } else {
      console.error(
        "Current subdomain is invalid; must be puddle.dgw.cloud, contoso.dgw.cloud, or xyz.dgw.cloud."
      );
    }

    return (
      <div className="App">
        <header className={headerClassName}>
          <p>
            Welcome to <code>{window.location.hostname}!</code>
          </p>
        </header>
        {nav}
        {actions}
      </div>
    );
  }

  private _redirect(target: string): void {
    window.location.replace(target);
  }

  private _getRange(): void {
    let range;
    Excel.run(async context => {
      range = context.workbook.getSelectedRange();
      range.load("address");

      await context.sync();

      console.log(range.address);

      this.setState({ range: range.address });
    }).catch(err => {
      console.log(err);
    });
  }
}

export default App;
