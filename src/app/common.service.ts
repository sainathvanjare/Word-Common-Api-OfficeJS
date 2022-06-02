import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { BehaviorSubject, Observable } from 'rxjs';
import { map } from 'rxjs/operators';


@Injectable({
  providedIn: 'root'
})
export class CommonService {
  private currentUserSubject
  public currentUser
  baseUrl = "https://officejs-server.herokuapp.com/api"

  constructor(private http: HttpClient) {
      this.currentUserSubject = (JSON.parse(localStorage.getItem('currentUser')));
      this.currentUser = this.currentUserSubject
  }

  public get currentUserValue() {
      return this.currentUserSubject.value;
  }

  login(username, password) {
    let data = {
      "user":
      {
        "email": username,
        "password": password
      }
    }
      return this.http.post<any>(`${this.baseUrl}/users/login`, data)
          .pipe(map(user => {
              // store user details and jwt token in local storage to keep user logged in between page refreshes
              localStorage.setItem('currentUser', JSON.stringify(user));
              return user;
          }, (error)=>{
            console.log(error)
          }));
  }

  logout() {
      // remove user from local storage and set current user to null
      localStorage.removeItem('currentUser');
      this.currentUserSubject.next(null);
  }
}
