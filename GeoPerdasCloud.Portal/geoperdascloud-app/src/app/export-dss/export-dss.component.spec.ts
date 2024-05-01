import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExportDssComponent } from './export-dss.component';

describe('ExportDssComponent', () => {
  let component: ExportDssComponent;
  let fixture: ComponentFixture<ExportDssComponent>;

  beforeEach(() => {
    TestBed.configureTestingModule({
      declarations: [ExportDssComponent]
    });
    fixture = TestBed.createComponent(ExportDssComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
