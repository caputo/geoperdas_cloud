import { TestBed } from '@angular/core/testing';

import { GeoperdasServiceService } from './geoperdas-service.service';

describe('GeoperdasServiceService', () => {
  let service: GeoperdasServiceService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(GeoperdasServiceService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
