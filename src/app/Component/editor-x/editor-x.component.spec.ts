import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EditorXComponent } from './editor-x.component';

describe('EditorXComponent', () => {
  let component: EditorXComponent;
  let fixture: ComponentFixture<EditorXComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EditorXComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EditorXComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
