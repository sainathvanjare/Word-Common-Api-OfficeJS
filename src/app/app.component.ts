import { Component, OnInit } from '@angular/core';
import { AppService } from "./app.service";
// import * as $ from 'jquery';
// import { range } from 'rxjs';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {

  constructor() {
  }
  title = 'word-addin';
  
}
