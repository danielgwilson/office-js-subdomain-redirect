import React, { Component } from "react";
import "./App.css";
import {
  DefaultButton,
  PrimaryButton
} from "office-ui-fabric-react/lib/Button";

class App extends Component<any, any> {
  constructor(props: any) {
    super(props);
    this.state = {
      range: "",
      abc: window.location.hostname == "abc.dgw.cloud"
    };
    this._redirect = this._redirect.bind(this);
    this._getRange = this._getRange.bind(this);
  }
  render() {
    return (
      <div className="App">
        <header className={this.state.abc ? "App-header" : "App-header-alt"}>
          <p>
            Welcome to <code>{window.location.hostname}!</code>
          </p>
        </header>
        <div>
          <p>
            Redirect to{" "}
            <code>{this.state.abc ? "xyz.dgw.cloud" : "abc.dgw.cloud"}</code>
          </p>
          <PrimaryButton style={{}} text="Redirect" onClick={this._redirect} />
        </div>
        <div>
          <p>Get the range of the current selection</p>
          <DefaultButton style={{}} text="Get Range" onClick={this._getRange} />
          <p>{this.state.range}</p>
        </div>
      </div>
    );
  }

  private _redirect(): void {
    window.location.replace(
      this.state.abc ? "https://xyz.dgw.cloud" : "https://abc.dgw.cloud"
    );
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
