import * as React from 'react';
import styles from './ImportCsv.module.scss';
import { IImportCsvProps } from './IImportCsvProps';
import { readRemoteFile, jsonToCSV } from 'react-papaparse';
import { IImportCSVState } from './IImportCSVState';
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { PrimaryButton } from 'office-ui-fabric-react';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const ExcelData: any[] = [];
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
          //#region Filter data and Print in Console
          console.log("Row data:", results.data);
         
          //#endregion
          
          ExcelData.push(results.data);
          this.setState({ ExcelFileData: ExcelData});
            
          //#region Export to CSV
          const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
          const fileExtension = '.xlsx';
          const ws = XLSX.utils.json_to_sheet(this.state.ExcelFileData);
          const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
          const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const data = new Blob([excelBuffer], { type: fileType });
          FileSaver.saveAs(data, "ABC" + fileExtension);
          //#endregion

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
