import { Component } from '@angular/core';
import {HttpClient, HttpHeaders} from '@angular/common/http'
import { GeoPerdasConfig } from '../model/config';
import { GeoperdasServiceService } from '../geoperdas-service.service';


@Component({
  selector: 'app-export-dss',
  templateUrl: './export-dss.component.html',
  styleUrls: ['./export-dss.component.css']
})
export class ExportDssComponent {
  constructor(private http:HttpClient, private geoperdasService:GeoperdasServiceService){}
  feeder!:string; 
  config:GeoPerdasConfig = new GeoPerdasConfig();
  loading:Boolean = false;
  errorMessage:String = "";
  export(){    
    this.loading = true;  

    this.geoperdasService.exportDss(this.config).subscribe((response) => {
      const url = window.URL.createObjectURL(response);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'file.zip';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);      
      this.loading = false;
    },(error)=>{
      console.log(error);
      this.errorMessage = error.message;
      this.loading = false;
      this.errorMessage="";
    }); 
  }

  requestCalc()
  {    
    this.loading = true;        
    this.geoperdasService.postRequest(this.config).subscribe((result) => {
      this.loading = false;
    },(error)=>{
      console.log(error);
      this.errorMessage = error.message;
      this.loading = false;
      setTimeout(() => {
        this.errorMessage="";
      }, 5000);
    });
  }
  
}
