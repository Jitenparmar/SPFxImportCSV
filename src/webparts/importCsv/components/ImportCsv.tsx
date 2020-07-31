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

/**
 * Git Hub Desktop changes
 */


const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const ExcelData: any[] = [];
var ws;
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
    this._getOutputTemplateFile();
  }
  public _getOutputTemplateFile = () => {
    type ColInfo = {
      wpx?: 50;  // width in screen pixels
    };

    ws = XLSX.utils.aoa_to_sheet([["Analysis No"]]);
    /* Set worksheet sheet to "normal" */
    ws["!margins"] = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
    var wscols = [
      { wch: 18 },
      { wch: 15 },
      { wch: 15 },
      { wch: 10 }
    ];
    ws["!cols"] = wscols;
    XLSX.utils.sheet_add_aoa(ws, [["CATIONS"]], { origin: "A3" });
    XLSX.utils.sheet_add_aoa(ws, [["Cations", "Result", "Units", "Rerun?"]], { origin: "B4" });


    XLSX.utils.sheet_add_aoa(ws, [["Calcium (Ca)"], ["Magnesium (Mg)"], ["Sodium (Na)"], ["Potassium (K)"], ["Iron (Fe)"], ["Manganese (Mn)"], ["Aluminum (Al)"], ["Barium (Ba)"], ["Strontium (Sr)"], ["Copper (Cu)"], ["Zinc (Zn)"]], { origin: "B5" });
    XLSX.utils.sheet_add_aoa(ws, [["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l"], ["mg/l"], ["mg/l"], ["mg/l"], ["mg/l"], ["mg/l"], ["mg/l"]], { origin: "D5" });

    XLSX.utils.sheet_add_aoa(ws, [["ANIONS"]], { origin: "A17" });
    XLSX.utils.sheet_add_aoa(ws, [["Anions", "Result", "Units", "Rerun?"]], { origin: "B18" });

    XLSX.utils.sheet_add_aoa(ws, [["Bicarb (HCO3)"], ["Fluoride (F)"], ["Chloride (Cl)"], ["Bromide (Br)"], ["Nitrate (NO3)"], ["Phosphate (PO4)"], ["Sulfate (SO4)"], ["Silica (SiO2)"]], { origin: "B19" });
    XLSX.utils.sheet_add_aoa(ws, [["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"], ["mg/l CaCO3"]], { origin: "D19" });

    XLSX.utils.sheet_add_aoa(ws, [["OTHER PARAMETERS"]], { origin: "A28" });
    XLSX.utils.sheet_add_aoa(ws, [["Parameter", "Result", "Units", "Rerun?"]], { origin: "B29" });

    XLSX.utils.sheet_add_aoa(ws, [["pH"], ["*Turbidity"], ["*Conductivity"], ["Total Hardness"], ["TOC (C)"], ["Free (CO2)"]], { origin: "B30" });
    XLSX.utils.sheet_add_aoa(ws, [["'--"], ["NTU"], ["uS/cm"], ["mg/l CaCO3"], ["mg/l"], ["mg/l CaCO3"]], { origin: "D30" });

    XLSX.utils.sheet_add_aoa(ws, [["WEIGHTS"]], { origin: "A37" });
    XLSX.utils.sheet_add_aoa(ws, [["Weight Type", "Gross", "Tare", "Units"]], { origin: "B38" });

    XLSX.utils.sheet_add_aoa(ws, [["*TSS"], ["*TDS"], ["*TS"]], { origin: "B39" });
    XLSX.utils.sheet_add_aoa(ws, [[10], [20]], { origin: "C5" });


    /**
     * 
      
    const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
    const fileExtension = '.xlsx';
    ws['C33'] = {f: 'SUM(C5:C6)'};
    const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const data = new Blob([excelBuffer], { type: fileType });
    FileSaver.saveAs(data, "ABCOutput" + fileExtension);
    */
  }
  public _getCSVFileData = () => {
    // console.log("Sample ID" + this.state.SampleID);
    readRemoteFile('https://evoqua.sharepoint.com/sites/IntranetQA/InstumentFiles/UV Spec.csv', {
      header: true,
      step: (results, parser) => {
        if (results.data["Sample ID"] == this.state.SampleID) {
          //#region Filter data and Print in Console
          console.log("Row data:", results.data["SiO2 as CaCO3"]);
          XLSX.utils.sheet_add_aoa(ws, [[results.data["SiO2 as CaCO3"]]], { origin: "C26" });
          readRemoteFile('https://evoqua.sharepoint.com/sites/IntranetQA/InstumentFiles/Metrohm Titrator.csv', {
            header: true,
            step: (results, parser) => {
              if (results.data["ID1.Value"] == this.state.SampleID) {
                console.log("Row data:", results.data["RS01.Value"]);

                XLSX.utils.sheet_add_aoa(ws, [[results.data["RS01.Value"]]], { origin: "C30" });
                XLSX.utils.sheet_add_aoa(ws, [[results.data["RS03.Value"]]], { origin: "C19" });

                readRemoteFile('https://evoqua.sharepoint.com/sites/IntranetQA/InstumentFiles/Metrohm IC.csv', {
                  header: true,
                  step: (results, parser) => {
                    if (results.data["Ident"] == this.state.SampleID) {

                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.F"]]], { origin: "C20" });
                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.NO3"]]], { origin: "C23" });
                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.PO4"]]], { origin: "C24" });
                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.SO4"]]], { origin: "C25" });
                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.Br"]]], { origin: "C22" });
                      XLSX.utils.sheet_add_aoa(ws, [[results.data["RS.Cl"]]], { origin: "C21" });
                      readRemoteFile('https://evoqua.sharepoint.com/sites/IntranetQA/InstumentFiles/P519730740 STD TOC.csv', {
                        header: true,
                        step: (results, parser) => {
                          if (results.data["Sample ID"] == this.state.SampleID) {

                            XLSX.utils.sheet_add_aoa(ws, [[results.data["Conc (PPM as mg/L C)"]]], { origin: "C34" });

                            const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
                            const fileExtension = '.xlsx';
                            ws['C33'] = { f: 'SUM(C5:C6)' };
                            const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
                            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                            const data = new Blob([excelBuffer], { type: fileType });
                            FileSaver.saveAs(data, "ABCOutput" + fileExtension);

                          }
                        }
                      });
                    }
                  }
                });



              }
            }
          });



          // ExcelData.push(results.data);
          // this.setState({ ExcelFileData: ExcelData });

          //#region Export to CSV
          /** 
          const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
          const fileExtension = '.xlsx';
          const ws = XLSX.utils.json_to_sheet(this.state.ExcelFileData);
          const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
          const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const data = new Blob([excelBuffer], { type: fileType });
          FileSaver.saveAs(data, "ABC" + fileExtension);*/
          //#endregion

        }
      }
    });


  }
  public _setTextboxValuetoState = (e: any) => {
    var data = e.target.value;
    XLSX.utils.sheet_add_aoa(ws, [[data]], { origin: "B1" });
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
