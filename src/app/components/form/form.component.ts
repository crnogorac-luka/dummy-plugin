import { Component, OnInit } from '@angular/core';
import {
  FormBuilder,
  FormControl,
  FormGroup,
  Validators,
} from '@angular/forms';
import { take, catchError } from 'rxjs/operators';
import { of, throwError } from 'rxjs';
import { Error } from 'src/app/model/Error';
import { Fruit } from 'src/app/model/Fruit';
import { ApiService } from 'src/app/services/api.service';
import { HttpErrorResponse } from '@angular/common/http';

@Component({
  selector: 'app-form',
  templateUrl: './form.component.html',
  styleUrls: ['./form.component.scss'],
})
export class FormComponent implements OnInit {
  form!: FormGroup;

  fruitObject: Fruit;

  isError: boolean = false;

  constructor(private apiService: ApiService, private fb: FormBuilder) {}

  ngOnInit() {
    this.form = this.fb.group({
      fruitName: ['', Validators.required],
    });
  }

  async getFruitData(form: FormGroup) {
    try {
      const fruit = form.controls['fruitName'].value.toLowerCase().trim();

      Office.onReady(() => {
        Excel.run(async (context: Excel.RequestContext) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();

          // Load worksheet properties
          sheet.load('name, usedRange');

          // Sync to apply loaded properties
          await context.sync();

          console.log('Worksheet name: ' + sheet.name);
          console.log('Used range: ' + sheet.getUsedRange());

          const range = sheet.getUsedRange();

          // Load range properties
          range.load('values');

          // Sync to apply loaded properties
          await context.sync();

          console.log('Used range values: ' + range.values);

          const usedRangeValues = range.values;
          if (
            !usedRangeValues ||
            usedRangeValues.length === 0 ||
            usedRangeValues.length === 1
          ) {
            const headerRow = sheet.getRange('B2:H2');
            headerRow.load('values,format');

            // Sync to apply loaded properties
            await context.sync();

            // Define the header row format
            headerRow.format.fill.color = 'red';
            headerRow.format.fill.tintAndShade = 0.5;

            // Set the header row values
            headerRow.values = [
              [
                'name',
                'family',
                'calories',
                'fat',
                'sugar',
                'carbohydrates',
                'protein',
              ],
            ];

            console.log('true and added');
          }

          // Now that previous operations are done, you can make the API call
          this.apiService
            .getFruit(fruit)
            .pipe(take(1))
            .subscribe(
              (fruitData: Fruit) => {
                this.fruitObject = new Fruit(
                  fruitData.name,
                  fruitData.family,
                  fruitData.nutritions.calories,
                  fruitData.nutritions.fat,
                  fruitData.nutritions.sugar,
                  fruitData.nutritions.carbohydrates,
                  fruitData.nutritions.protein
                );
                console.log('enter subscribe');
                this.populateRow(this.fruitObject);
                this.isError = false;
              },
              (error: HttpErrorResponse) => {
                this.isError = true;
                console.log('test', error.error.error);
              }
            );

          // Sync one last time to ensure all changes are applied
          await context.sync();
        });
      });
    } catch (error) {
      console.error(`Error: ${error}`);
      console.log('test');
    }
  }

  async populateRow(fruit: Fruit) {
    Office.onReady(() => {
      Excel.run(async (context: Excel.RequestContext) => {
        // Get a reference to the active worksheet
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Define the row where you want to populate the data (e.g., row 3)
        const rowNumber = 2;

        // Create an array with the values you want to populate
        const valuesToPopulate = [
          fruit.name,
          fruit.family,
          fruit.nutritions.calories,
          fruit.nutritions.fat,
          fruit.nutritions.sugar,
          fruit.nutritions.carbohydrates,
          fruit.nutritions.protein,
        ];

        // Get a reference to the range starting from column B (second column) in the specified row
        const range = sheet.getRangeByIndexes(
          rowNumber,
          1,
          1,
          valuesToPopulate.length
        );

        // Load the range
        range.load('values');

        // Sync to apply the loaded properties
        await context.sync();

        // Set the values of the range with the values from the Fruit object
        range.values = [valuesToPopulate];

        // Sync again to apply the changes
        await context.sync();

        console.log('Row 3 has been populated with Fruit object properties.');
      });
    });
  }
}
