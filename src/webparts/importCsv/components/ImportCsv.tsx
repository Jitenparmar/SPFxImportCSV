import * as React from 'react';
import styles from './ImportCsv.module.scss';
import { IImportCsvProps } from './IImportCsvProps';
import { readRemoteFile } from 'react-papaparse';
import { IImportCSVState } from './IImportCSVState';
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { PrimaryButton } from 'office-ui-fabric-react';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};
export default class ImportCsv extends React.Component<IImportCsvProps, IImportCSVState> {
  constructor(props: IImportCsvProps) {
    super(props);
    this.state = {
      Loading: false,
      ExcelFileData: [],
      SampleID: null
    };
  }

  public _getCSVFileData = () => {
    console.log("Sample ID" + this.state.SampleID);
    readRemoteFile('https://evoqua.sharepoint.com/sites/IntranetQA/InstumentFiles/H013730328_DI.csv', {
      header: true,
      step: (results, parser) => {
        if (results.data["Sample ID"] == this.state.SampleID) {
          console.log("Row data:", results.data);
          this.setState({ ExcelFileData: results.data });
        }
      }
    });
  }
  public _setTextboxValuetoState = (e: any) => {
    var data = e.target.value;
    this.setState({ SampleID: data });
  }
  public render(): React.ReactElement<IImportCsvProps> {
    return (
      <Stack>
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack.Item grow>
            <TextField label="Sample ID:" underlined onBlur={this._setTextboxValuetoState} />
          </Stack.Item>
          <PrimaryButton
            text="Submit"
            onClick={this._getCSVFileData}
            allowDisabledFocus />
        </Stack>
      </Stack>
    );
  }
}
