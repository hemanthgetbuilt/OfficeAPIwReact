import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [],
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  createSheet = async () => {
    try {
      await Excel.run(async (context) => {
        context.workbook.worksheets.add("Test Sheet");
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  enterdatainrange = async () => {
    try {
      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Test Sheet");
        let data = [
          ["Test Cell#11", "Test Cell#12", 13],
          ["Test Cell#21", "Test Cell#22", 23],
          ["Test Cell#31", "Test Cell#32", 33],
        ];
        let range = sheet.getRange("B5:D7");
        range.values = data;
        let formulaRange = sheet.getRange("D8");
        formulaRange.formulas = [["=SUM(D5:D7)"]];
        range.format.autofitColumns();
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  alert = async () => {
    await Excel.run(async (context) => {
      Office.context.ui.displayDialogAsync("https://localhost:3000/dialogbox.html", {
        height: 30,
        width: 20,
        displayInIframe: true,
      });
      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/built_min_logo.png")} title={this.props.title} message="Welcome" />
        <HeroList message="" items={this.state.listItems}>
          <p className="ms-font-l">Modify the source files, then try.</p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Highlight Cell
          </DefaultButton>
          <br />
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.createSheet}
          >
            Create New Sheet
          </DefaultButton>
          <br />
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.enterdatainrange}
          >
            Enter Data
          </DefaultButton>
          <br />
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.alert}>
            Show Dialog
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
