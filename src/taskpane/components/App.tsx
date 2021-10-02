import * as React from "react";
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DetailsList, SelectionMode, CheckboxVisibility } from '@fluentui/react/lib/DetailsList';
import Progress from "./Progress";
/* global Excel  */

export interface AppProps {
  isOfficeInitialized: boolean;
}

export interface AppState {
  sheetInfos: SheetInfo[];
  filterWord: string;
}

interface SheetInfo {
  id: string;
  name: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      sheetInfos: [],
      filterWord: null,
    };
  }

  componentDidMount() {
    this.registerEventHandlers();
    this.resetSheetNames();
  }

  registerEventHandlers = async () => {
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.onAdded.add(this.resetSheetNames);
      worksheets.onDeleted.add(this.resetSheetNames);
    });
  };

  resetSheetNames = async () => {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items");
      await context.sync();

      const sheetInfos = sheets.items
        .filter((sheet) => sheet.visibility === Excel.SheetVisibility.visible)
        .filter((sheet) => {
          if (!this.state.filterWord) return true;

          const regex = new RegExp(this.state.filterWord, "i"); // ignore case
          return regex.test(sheet.name);
        })
        .map((sheet) => {
          return {
            id: sheet.id,
            name: sheet.name,
          };
        });

      this.setState({ sheetInfos });
    });
  };

  onFilterWordChange = (_: React.ChangeEvent, value: string) => {
    this.setState({
      filterWord: value,
    });
    this.resetSheetNames();
  };

  onActiveItemChanged = async (sheetInfo: SheetInfo) => {
    await Excel.run(async (context) => {
      try {
        const sheetId = sheetInfo.id;
        const sheet = context.workbook.worksheets.getItem(sheetId);
        sheet.load("name");
        sheet.activate();
        await context.sync();
        if (sheet.name !== sheetInfo.name) return this.resetSheetNames();
      } catch {
        return this.resetSheetNames();
      }
    });
  };

  render() {
    const { isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return <Progress message="Initializing..." />;
    }

    return (
      <div className="hdobysh-exaddin">
        <SearchBox
          onChange={this.onFilterWordChange}
          placeholder="Filter"
          iconProps={{ iconName: "Filter" }}
          showIcon={true}
        />
        <DetailsList
          className="hdobysh-exaddin__detailslist"
          items={this.state.sheetInfos}
          selectionMode={SelectionMode.single}
          selectionPreservedOnEmptyClick={true}
          isHeaderVisible={false}
          checkboxVisibility={CheckboxVisibility.hidden}
          columns={[{ key: "name", name: "name", fieldName: "name", minWidth: 100 }]}
          onActiveItemChanged={this.onActiveItemChanged}
        />
      </div>
    );
  }
}
