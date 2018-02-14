import { Component } from '@angular/core';
import * as XLSX from 'ts-xlsx';
import {isNullOrUndefined} from "util";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  arrayBuffer: any;
  file: File;
  output: any;
  sheetNumber: number = 0;

  incomingFile(event) {
    this.file = event.target.files[0];
  }

  Upload() {
    if (!this.file) {
      alert("Please select xmls file.");
      return false;
    }

    let fileReader = new FileReader();
    fileReader.onload = (e) => {
      this.arrayBuffer = fileReader.result;
      let data = new Uint8Array(this.arrayBuffer);
      let arr = new Array();

      for(let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);

      let bstr = arr.join("");
      let workbook = XLSX.read(bstr, {type:"binary"});

      if (this.sheetNumber === null || this.sheetNumber === undefined || this.sheetNumber < 0) {
        this.sheetNumber = 0;
      }

      let sheet_name = workbook.SheetNames[this.sheetNumber];
      let worksheet = workbook.Sheets[sheet_name];

      // console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}));
      this.output = JSON.stringify(XLSX.utils.sheet_to_json(worksheet,{raw:true}), null, 4);
    }
    fileReader.readAsArrayBuffer(this.file);
  }
}
