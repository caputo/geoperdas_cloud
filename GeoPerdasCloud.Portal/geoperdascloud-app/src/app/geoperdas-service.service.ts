import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { GeoPerdasConfig } from './model/config';
import { Observable } from 'rxjs';
import { isDevMode } from '@angular/core';
@Injectable({
  providedIn: 'root'
})
export class GeoperdasServiceService {

  constructor(private http:HttpClient) { }

  private readonly exportDssPath:string = `${isDevMode()?"https://localhost:44387":""}/ExportDss/`;
  private readonly requestCalcPath:string = `${isDevMode()?"https://localhost:44387":""}/Request/`;


  exportDss(config: GeoPerdasConfig):Observable<Blob> {    
    let headers = new HttpHeaders().set('Content-Type', 'application/json');
    headers.set('Content-Type', 'application/blob');
    const body = config;
    return this.http.post(this.exportDssPath, body, { responseType: 'blob' })    
  }

  postRequest(config: GeoPerdasConfig):Observable<Object>
  { 
    const body = config;
    const headers = new HttpHeaders().set('Content-Type', 'application/json');    
    return this.http.post(this.requestCalcPath, body);
  }


}
