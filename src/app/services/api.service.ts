import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';
import { Fruit } from '../model/Fruit';

@Injectable({
  providedIn: 'root',
})
export class ApiService {
  private apiUrl = 'https://cors-anywhere.herokuapp.com/https://www.fruityvice.com/api/fruit';

  httpOptions = {
    headers: new HttpHeaders({
      'Content-Type': 'application/json',
    }),
  };

  constructor(private http: HttpClient) {}

  getFruit(fruit: string): Observable<Fruit> {
    const url = `${this.apiUrl}/${fruit}`;
    return this.http.get<Fruit>(url);
  }
}
