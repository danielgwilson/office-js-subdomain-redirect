import React, { Component } from "react";
import "./App.css";

class App extends Component {
    render() {
        return (
            <div className="App">
                <header className="App-header">
                    <p>
                        Welcome to <code>xyz.dgw.cloud!</code>
                    </p>
                    <a
                        className="App-link"
                        href="https://abc.dgw.cloud"
                        target="_self"
                        rel="noopener noreferrer">
                        Redirect to <code>abc.dgw.cloud</code>
                    </a>
                </header>
            </div>
        );
    }
}

export default App;
