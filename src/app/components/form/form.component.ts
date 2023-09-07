import { Component, OnInit } from '@angular/core';
import { NgForm } from '@angular/forms';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from 'src/app/app.module';

@Component({
  selector: 'app-form',
  templateUrl: './form.component.html',
  styleUrls: ['./form.component.scss'],
})
export class FormComponent implements OnInit {

  number: any;

  constructor() {}

  ngOnInit(): void {
    Office.initialize = function () {
      const platform = platformBrowserDynamic();
      platform.bootstrapModule(AppModule);
    };
  }

  async changeColor() {
    const cellIdentifier = this.number;
    console.log(cellIdentifier);
    try {
      // Parse the cell identifier to extract row and column information.
      const match = cellIdentifier.match(/([A-Z]+)(\d+)/);
  
      if (!match) {
        console.error(`Invalid cell identifier: ${cellIdentifier}`);
        return;
      }
  
      const columnName = match[1];
      const rowNumber = parseInt(match[2]);
      Office.onReady(() => {
        Excel.run(async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cell = sheet.getRange(cellIdentifier);
  
        // You can modify the cell's value or format here if needed.
        cell.format.fill.color = "blue";

        await context.sync();
        
      });
    });
    } catch (error) {
      console.error(`Error: ${error}`);
    }
  }

  
}
