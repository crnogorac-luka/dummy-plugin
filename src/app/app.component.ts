import { Component, OnInit } from '@angular/core';
import { AppModule } from './app.module';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit{
  title = 'dummy-plugin';

  ngOnInit(): void {
    Office.initialize = function () {
      const platform = platformBrowserDynamic();
      platform.bootstrapModule(AppModule);
    };
  }
}
