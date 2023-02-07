import * as XLSX from 'xlsx';
import { Subject } from "rxjs";
import { Component, ElementRef, ViewChild } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})

export class AppComponent {

  @ViewChild('inputFile') inputFile!: ElementRef;

  dataSheet: any = new Subject();
  keys: string[] | null = [];
  IsNotExcelFile: boolean = false;
  spinnerEnabled: boolean = false;

  onChange(evt: any) {
    let data: any;
    const target: DataTransfer = <DataTransfer>(evt.target);
    this.IsNotExcelFile = !target.files[0].name.match(/(.xls|.xlsx)/);
    if (target.files.length > 1) { this.inputFile.nativeElement.value = ''; }
    if (!this.IsNotExcelFile) {
      this.spinnerEnabled = true;
      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        /* read workbook */
        const bstr: string = e.target.result;
        const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

        /* grab first sheet */
        const wsname: string = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];

        /* save data */
        data = XLSX.utils.sheet_to_json(ws);
      };
      reader.readAsBinaryString(target.files[0]);
      reader.onloadend = (e) => {
        this.spinnerEnabled = false;
        this.keys = Object.keys(data[0]);
        this.dataSheet.next(data)
      }
    } else { this.inputFile.nativeElement.value = ''; }
  }

  removeData() {
    this.inputFile.nativeElement.value = '';
    this.dataSheet.next(null);
    this.keys = null;
  }

}
