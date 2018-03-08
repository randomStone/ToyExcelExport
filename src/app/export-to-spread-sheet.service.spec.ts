import { TestBed, inject } from '@angular/core/testing';

import { ExportToSpreadSheetService } from './export-to-spread-sheet.service';

describe('ExportToSpreadSheetService', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [ExportToSpreadSheetService]
    });
  });

  it('should be created', inject([ExportToSpreadSheetService], (service: ExportToSpreadSheetService) => {
    expect(service).toBeTruthy();
  }));
});
