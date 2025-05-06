import { Component, ElementRef, ViewChild, ViewChildren, QueryList, AfterViewInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

interface Margins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  // Border settings
  showBorder: boolean;
  borderWidth: number;
  borderStyle: string;
  borderColor: string;
  // Individual border settings
  borderTopWidth: number;
  borderBottomWidth: number; 
  borderLeftWidth: number;
  borderRightWidth: number;
  borderTopStyle: string;
  borderBottomStyle: string;
  borderLeftStyle: string;
  borderRightStyle: string;
  borderTopColor: string;
  borderBottomColor: string;
  borderLeftColor: string;
  borderRightColor: string;
}

@Component({
  selector: 'app-editor-x',
  templateUrl: './editor-x.component.html',
  styleUrl: './editor-x.component.css',
  imports: [CommonModule, FormsModule],
  standalone: true
})
export class EditorXComponent implements AfterViewInit {
  @ViewChildren('editorCanvas') editorCanvasList!: QueryList<ElementRef>;

  @ViewChild('fileInput') fileInput!: ElementRef;

  // Image Modal Properties
  showImageModal: boolean = false;
  imageFile: File | null = null;
  imageUrl: string = '';
  imageWidth: number | null = null;
  imageHeight: number | null = null;

  // Shape Modal Properties
  showShapeModal: boolean = false;
  selectedShape: 'rectangle' | 'circle' | 'triangle' | null = null;
  shapeWidth: number = 100;
  shapeHeight: number = 100;
  shapeFillColor: string = '#f0f0f0';
  shapeBorderColor: string = '#999999';

  // Line Modal Properties
  showLineModal: boolean = false;
  selectedLine: 'horizontal' | 'vertical' | null = null;
  lineLength: number = 200;
  lineThickness: number = 2;
  lineColor: string = '#000000';

  // Element Management
  private selectedElement: HTMLElement | null = null;
  
  currentDate: string = new Date().toLocaleDateString('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  });
  
  // Array to hold multiple pages
  pages: number[] = [0];
  
  // Zoom control
  zoomLevel: number = 100;
  minZoom: number = 50;
  maxZoom: number = 200;
  zoomStep: number = 10;
  
  // Modal controls
  showExportModal: boolean = false;
  showFieldsModal: boolean = false;
  showTableModal: boolean = false;
  showMarginModal: boolean = false;
  
  // Tab control for margin modal
  activeTab: string = 'margins-tab';
  
  // Table grid selection
  maxGridRows: number = 8;
  maxGridCols: number = 8;
  gridRows: number = 1;
  gridCols: number = 1;
  
  // Margin and border controls
  margins: Margins = {
    top: 2.0,
    bottom: 2.0,
    left: 2.0,
    right: 2.0,
    // Single border settings
    showBorder: false,
    borderWidth: 1,
    borderStyle: 'solid',
    borderColor: '#000000',
    // Individual border settings
    borderTopWidth: 1,
    borderBottomWidth: 1,
    borderLeftWidth: 1,
    borderRightWidth: 1,
    borderTopStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderRightStyle: 'solid',
    borderTopColor: '#000000',
    borderBottomColor: '#000000',
    borderLeftColor: '#000000',
    borderRightColor: '#000000'
  };
  
  defaultMargins: Margins = {
    top: 2.0,
    bottom: 2.0,
    left: 2.0,
    right: 2.0,
    // Single border settings
    showBorder: false,
    borderWidth: 1,
    borderStyle: 'solid',
    borderColor: '#000000',
    // Individual border settings
    borderTopWidth: 1,
    borderBottomWidth: 1,
    borderLeftWidth: 1,
    borderRightWidth: 1,
    borderTopStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderRightStyle: 'solid',
    borderTopColor: '#000000',
    borderBottomColor: '#000000',
    borderLeftColor: '#000000',
    borderRightColor: '#000000'
  };

  // Border style options
  borderStyles = [
    { value: 'solid', label: 'Solid' },
    { value: 'dashed', label: 'Dashed' },
    { value: 'dotted', label: 'Dotted' },
    { value: 'double', label: 'Double' }
  ];
  
  // Store selection range for field insertion
  private lastSelection: Range | null = null;
  
  ngAfterViewInit(): void {
    // Set focus to the first editor canvas after view initialization
    setTimeout(() => {
      if (this.editorCanvasList.first) {
        this.editorCanvasList.first.nativeElement.focus();
        
        // Apply initial margins and border to all pages
        this.applyMargins();
      }
      
      // Initialize table resizing for any existing tables
      this.makeTablesResizable();
      this.makeTablesDraggable();
      
      // Set up observer to handle dynamically added tables
      this.observeContentChanges();
    });
  
    setTimeout(() => {
      this.initializeDraggableElements();
    }, 500);
    
    // Enhanced keyboard event listeners for element deletion
    document.addEventListener('keydown', (e) => {
      if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedElement) {
        // Prevent default action for backspace to avoid browser navigation
        if (e.key === 'Backspace') {
          e.preventDefault();
        }
        
        // Make sure we're not inside a contenteditable element
        const activeElement = document.activeElement;
        if (activeElement && 
            (activeElement.getAttribute('contenteditable') === 'true' || 
             activeElement.tagName === 'INPUT' || 
             activeElement.tagName === 'TEXTAREA')) {
          // Allow normal deletion within editable elements
          return;
        }
        
        this.deleteSelectedElement();
      }
    });
  }

  /**
   * Delete the currently selected element
   */
  deleteSelectedElement(): void {
    if (this.selectedElement) {
      // If it's a table, or any other editor element
      if (this.selectedElement.tagName === 'TABLE' || 
          this.selectedElement.classList.contains('editor-element') ||
          this.selectedElement.classList.contains('editor-table-wrapper')) {
        
        // If it's a table wrapper, handle specially
        if (this.selectedElement.classList.contains('editor-table-wrapper')) {
          this.selectedElement.remove();
        } 
        // Otherwise remove the element directly
        else {
          this.selectedElement.remove();
        }
        
        // Reset the selected element reference
        this.selectedElement = null;
        
        // Focus back on the editor
        this.focusCurrentEditorCanvas();
      }
    }
  }

  /**
   * Open the image modal and save current selection
   */
  openImageModal(): void {
    // Save the current selection range
    this.saveCurrentSelection();
    
    // Reset image properties
    this.imageFile = null;
    this.imageUrl = '';
    this.imageWidth = null;
    this.imageHeight = null;
    
    this.showImageModal = true;
  }

  /**
   * Close the image modal
   */
  closeImageModal(): void {
    this.showImageModal = false;
    this.focusCurrentEditorCanvas();
  }

  /**
   * Trigger the file input click event
   */
  triggerFileInput(): void {
    this.fileInput.nativeElement.click();
  }

  /**
   * Handle image file upload
   * @param event The file input change event
   */
  handleImageUpload(event: Event): void {
    const input = event.target as HTMLInputElement;
    
    if (input.files && input.files.length > 0) {
      this.imageFile = input.files[0];
      
      // Preview image and get dimensions (optional)
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          // Default dimensions to actual image size if not set
          if (!this.imageWidth) this.imageWidth = img.width;
          if (!this.imageHeight) this.imageHeight = img.height;
        };
        img.src = e.target?.result as string;
      };
      reader.readAsDataURL(this.imageFile);
    }
  }

  /**
   * Insert an image at the cursor position
   */
  insertImage(): void {
    // If we have a saved selection, restore it
    if (this.lastSelection) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(this.lastSelection);
        
        let imgSrc = '';
        
        // Use file or URL based on input
        if (this.imageFile) {
          imgSrc = URL.createObjectURL(this.imageFile);
        } else if (this.imageUrl) {
          imgSrc = this.imageUrl;
        } else {
          // No image source provided
          return;
        }
        
        // Create the image element with enhanced resize controls (removed delete button)
        const imgHtml = `
        <div class="editor-element editor-image" contenteditable="false">
          <img src="${imgSrc}" style="${this.imageWidth ? 'width:' + this.imageWidth + 'px;' : ''} ${this.imageHeight ? 'height:' + this.imageHeight + 'px;' : ''}" alt="Inserted image">
          
          <!-- Corner handles -->
          <div class="resize-handle top-left"></div>
          <div class="resize-handle top-right"></div>
          <div class="resize-handle bottom-left"></div>
          <div class="resize-handle bottom-right"></div>
          
          <!-- Edge handles -->
          <div class="resize-handle top-center"></div>
          <div class="resize-handle bottom-center"></div>
          <div class="resize-handle middle-left"></div>
          <div class="resize-handle middle-right"></div>
          
          <!-- Rotation handle -->
          <div class="rotate-line"></div>
          <div class="rotate-handle"></div>
        </div>
        <p></p>
        `;
        
        // Insert a line break before the image if not at the beginning of a line
        selection.deleteFromDocument();
        document.execCommand('insertHTML', false, '<br>' + imgHtml);
        
        // Initialize drag and resize functionality after a small delay
        setTimeout(() => {
          this.initializeDraggableElements();
        }, 100);
      }
    }
    
    // Close the modal and focus back on the editor
    this.closeImageModal();
  }

  /**
   * Open the shape modal and save current selection
   */
  openShapeModal(): void {
    // Save the current selection range
    this.saveCurrentSelection();
    
    // Reset shape properties
    this.selectedShape = null;
    this.shapeWidth = 100;
    this.shapeHeight = 100;
    this.shapeFillColor = '#f0f0f0';
    this.shapeBorderColor = '#999999';
    
    this.showShapeModal = true;
  }

  /**
   * Close the shape modal
   */
  closeShapeModal(): void {
    this.showShapeModal = false;
    this.focusCurrentEditorCanvas();
  }

  /**
   * Select a shape type
   * @param shapeType The type of shape to select
   */
  selectShape(shapeType: 'rectangle' | 'circle' | 'triangle'): void {
    this.selectedShape = shapeType;
  }

  /**
   * Insert a shape at the cursor position
   */
  insertShape(): void {
    // If we have a saved selection, restore it
    if (this.lastSelection && this.selectedShape) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(this.lastSelection);
        
        let shapeHtml = '';
        
        // Create different shapes based on selection
        switch (this.selectedShape) {
          case 'rectangle':
            shapeHtml = `
            <div class="editor-element editor-shape shape-rectangle" contenteditable="false" style="width: ${this.shapeWidth}px; height: ${this.shapeHeight}px; background-color: ${this.shapeFillColor}; border: 1px solid ${this.shapeBorderColor};">
              <!-- Corner handles -->
              <div class="resize-handle top-left"></div>
              <div class="resize-handle top-right"></div>
              <div class="resize-handle bottom-left"></div>
              <div class="resize-handle bottom-right"></div>
              
              <!-- Edge handles -->
              <div class="resize-handle top-center"></div>
              <div class="resize-handle bottom-center"></div>
              <div class="resize-handle middle-left"></div>
              <div class="resize-handle middle-right"></div>
              
              <!-- Rotation handle -->
              <div class="rotate-line"></div>
              <div class="rotate-handle"></div>
            </div>
            <p></p>
            `;
            break;
            
          case 'circle':
            shapeHtml = `
            <div class="editor-element editor-shape shape-circle" contenteditable="false" style="width: ${this.shapeWidth}px; height: ${this.shapeHeight}px; background-color: ${this.shapeFillColor}; border: 1px solid ${this.shapeBorderColor};">
              <!-- Corner handles -->
              <div class="resize-handle top-left"></div>
              <div class="resize-handle top-right"></div>
              <div class="resize-handle bottom-left"></div>
              <div class="resize-handle bottom-right"></div>
              
              <!-- Edge handles -->
              <div class="resize-handle top-center"></div>
              <div class="resize-handle bottom-center"></div>
              <div class="resize-handle middle-left"></div>
              <div class="resize-handle middle-right"></div>
              
              <!-- Rotation handle -->
              <div class="rotate-line"></div>
              <div class="rotate-handle"></div>
            </div>
            <p></p>
            `;
            break;
            
          case 'triangle':
            // For triangle we need to adjust the border values based on width/height
            const halfWidth = Math.round(this.shapeWidth / 2);
            shapeHtml = `
            <div class="editor-element editor-shape" contenteditable="false" style="width: ${this.shapeWidth}px; height: ${this.shapeHeight}px; position: relative;">
              <div style="width: 0; height: 0; position: absolute; top: 0; left: 0; border-left: ${halfWidth}px solid transparent; border-right: ${halfWidth}px solid transparent; border-bottom: ${this.shapeHeight}px solid ${this.shapeFillColor};"></div>
              
              <!-- Corner handles -->
              <div class="resize-handle top-left"></div>
              <div class="resize-handle top-right"></div>
              <div class="resize-handle bottom-left"></div>
              <div class="resize-handle bottom-right"></div>
              
              <!-- Edge handles -->
              <div class="resize-handle top-center"></div>
              <div class="resize-handle bottom-center"></div>
              <div class="resize-handle middle-left"></div>
              <div class="resize-handle middle-right"></div>
              
              <!-- Rotation handle -->
              <div class="rotate-line"></div>
              <div class="rotate-handle"></div>
            </div>
            <p></p>
            `;
            break;
        }
        
        // Insert a line break before the shape if not at the beginning of a line
        selection.deleteFromDocument();
        document.execCommand('insertHTML', false, '<br>' + shapeHtml);
        
        // Initialize drag and resize functionality after a small delay
        setTimeout(() => {
          this.initializeDraggableElements();
        }, 100);
      }
    }
    
    // Close the modal and focus back on the editor
    this.closeShapeModal();
  }

  /**
   * Open the line modal and save current selection
   */
  openLineModal(): void {
    // Save the current selection range
    this.saveCurrentSelection();
    
    // Reset line properties
    this.selectedLine = null;
    this.lineLength = 200;
    this.lineThickness = 2;
    this.lineColor = '#000000';
    
    this.showLineModal = true;
  }

  /**
   * Close the line modal
   */
  closeLineModal(): void {
    this.showLineModal = false;
    this.focusCurrentEditorCanvas();
  }

  /**
   * Select a line orientation
   * @param lineType The type of line to select
   */
  selectLine(lineType: 'horizontal' | 'vertical'): void {
    this.selectedLine = lineType;
  }

  /**
   * Insert a line at the cursor position
   */
  insertLine(): void {
    // If we have a saved selection, restore it
    if (this.lastSelection && this.selectedLine) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(this.lastSelection);
        
        let lineHtml = '';
        
        // Create different lines based on selection
        if (this.selectedLine === 'horizontal') {
          lineHtml = `
          <div class="editor-element editor-line line-horizontal" contenteditable="false" style="width: ${this.lineLength}px; height: ${this.lineThickness}px; background-color: ${this.lineColor};">
            <!-- Corner handles -->
            <div class="resize-handle top-left"></div>
            <div class="resize-handle top-right"></div>
            <div class="resize-handle bottom-left"></div>
            <div class="resize-handle bottom-right"></div>
            
            <!-- Edge handles -->
            <div class="resize-handle top-center"></div>
            <div class="resize-handle bottom-center"></div>
            <div class="resize-handle middle-left"></div>
            <div class="resize-handle middle-right"></div>
            
            <!-- Rotation handle -->
            <div class="rotate-line"></div>
            <div class="rotate-handle"></div>
          </div>
          <p></p>
          `;
        } else {
          lineHtml = `
          <div class="editor-element editor-line line-vertical" contenteditable="false" style="width: ${this.lineThickness}px; height: ${this.lineLength}px; background-color: ${this.lineColor};">
            <!-- Corner handles -->
            <div class="resize-handle top-left"></div>
            <div class="resize-handle top-right"></div>
            <div class="resize-handle bottom-left"></div>
            <div class="resize-handle bottom-right"></div>
            
            <!-- Edge handles -->
            <div class="resize-handle top-center"></div>
            <div class="resize-handle bottom-center"></div>
            <div class="resize-handle middle-left"></div>
            <div class="resize-handle middle-right"></div>
            
            <!-- Rotation handle -->
            <div class="rotate-line"></div>
            <div class="rotate-handle"></div>
          </div>
          <p></p>
          `;
        }
        
        // Insert a line break before the line if not at the beginning of a line
        selection.deleteFromDocument();
        document.execCommand('insertHTML', false, '<br>' + lineHtml);
        
        // Initialize drag and resize functionality after a small delay
        setTimeout(() => {
          this.initializeDraggableElements();
        }, 100);
      }
    }
    
    // Close the modal and focus back on the editor
    this.closeLineModal();
  }

  /**
   * Initialize draggable and resizable functionality for all editor elements
   */
  private initializeDraggableElements(): void {
    setTimeout(() => {
      const elements = document.querySelectorAll('.editor-element');
      
      elements.forEach(element => {
        const el = element as HTMLElement;
        
        // Ensure proper positioning for draggable elements
        if (window.getComputedStyle(el).position !== 'absolute') {
          el.style.position = 'absolute';
          
          // Get current position instead of using left/top which might be unset
          const rect = el.getBoundingClientRect();
          const parentRect = el.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
          
          el.style.left = `${rect.left - parentRect.left}px`;
          el.style.top = `${rect.top - parentRect.top}px`;
        }
        
        // Skip elements that have already been initialized
        if (el.dataset['initialized'] === 'true') {
          return;
        }
        
        // Mark as initialized to avoid duplicate event listeners
        el.dataset['initialized'] = 'true';
        
        // Add click event to select the element
        el.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          
          // Deselect any previously selected element
          if (this.selectedElement) {
            this.selectedElement.classList.remove('selected');
          }
          
          // Select this element
          el.classList.add('selected');
          this.selectedElement = el;
        });
        
        // Make element draggable
        this.addDragFunctionality(el);
        
        // Add resize functionality using all the resize handles
        const resizeHandles = el.querySelectorAll('.resize-handle, .rotate-handle');
        resizeHandles.forEach(handle => {
          this.addResizeFunctionality(el, handle as HTMLElement);
        });
      });
      
      // Add a click event listener to the document to deselect elements
      document.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (!target.closest('.editor-element') && this.selectedElement) {
          this.selectedElement.classList.remove('selected');
          this.selectedElement = null;
        }
      });
    }, 300); // Increased timeout for DOM to be fully rendered
  }

  /**
   * Add drag functionality to an element
   * @param element The element to make draggable
   */
  private addDragFunctionality(element: HTMLElement): void {

    const computedStyle = window.getComputedStyle(element);
    if (computedStyle.position !== 'absolute') {
      const rect = element.getBoundingClientRect();
      const parentRect = element.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
      
      element.style.position = 'absolute';
      element.style.left = `${rect.left - parentRect.left}px`;
      element.style.top = `${rect.top - parentRect.top}px`;
      element.style.width = `${rect.width}px`;
      element.style.height = `${rect.height}px`;
    }
    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;
    
    const startDrag = (e: MouseEvent) => {
      // Don't start dragging if clicked on a resize handle or delete button
      if ((e.target as HTMLElement).classList.contains('resize-handle') || 
          (e.target as HTMLElement).classList.contains('rotate-handle') ||
          (e.target as HTMLElement).classList.contains('delete-element-button')) {
        return;
      }
      
      // Select the element if not already selected
      if (!element.classList.contains('selected')) {
        // Deselect any previously selected element
        if (this.selectedElement) {
          this.selectedElement.classList.remove('selected');
        }
        
        // Select this element
        element.classList.add('selected');
        this.selectedElement = element;
      }
      
      e.preventDefault();
      e.stopPropagation();
      
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      
      // Get starting position from computed styles
      const computedStyle = window.getComputedStyle(element);
      startLeft = parseInt(computedStyle.left, 10) || 0;
      startTop = parseInt(computedStyle.top, 10) || 0;
      
      // Set absolute positioning if not already set
      if (computedStyle.position !== 'absolute') {
        const rect = element.getBoundingClientRect();
        const parentRect = element.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
        
        element.style.position = 'absolute';
        element.style.left = `${rect.left - parentRect.left}px`;
        element.style.top = `${rect.top - parentRect.top}px`;
        element.style.width = `${rect.width}px`;
        element.style.height = `${rect.height}px`;
      }
      
      // Add event listeners for drag
      document.addEventListener('mousemove', drag);
      document.addEventListener('mouseup', endDrag);
      
      // Add a class to indicate dragging
      element.classList.add('dragging');
    };
    
    const drag = (e: MouseEvent) => {
      if (!isDragging) return;
      
      e.preventDefault();
      
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      
      element.style.left = `${startLeft + dx}px`;
      element.style.top = `${startTop + dy}px`;
    };
    
    const endDrag = (e: MouseEvent) => {
      if (!isDragging) return;
      
      e.preventDefault();
      
      isDragging = false;
      document.removeEventListener('mousemove', drag);
      document.removeEventListener('mouseup', endDrag);
      
      // Remove the dragging class
      element.classList.remove('dragging');
    };
    
    // Add mousedown event listener to the element itself for dragging
    element.addEventListener('mousedown', startDrag);
  }

  /**
   * Fix element position and visibility
   */
  private fixElementsPositionAndVisibility(): void {
    // Add a small delay to ensure DOM is ready
    setTimeout(() => {
      // Find all editor elements
      const elements = document.querySelectorAll('.editor-element');
      
      elements.forEach(el => {
        const element = el as HTMLElement;
        
        // Ensure proper selection event
        element.addEventListener('click', (e) => {
          e.stopPropagation();
          
          // Deselect previous elements
          document.querySelectorAll('.editor-element.selected').forEach(selected => {
            selected.classList.remove('selected');
          });
          
          // Select current element
          element.classList.add('selected');
          this.selectedElement = element;
        });
        
        // Fix element positioning
        if (window.getComputedStyle(element).position !== 'absolute') {
          element.style.position = 'absolute';
          const rect = element.getBoundingClientRect();
          const parentRect = element.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
          
          element.style.left = `${rect.left - parentRect.left}px`;
          element.style.top = `${rect.top - parentRect.top}px`;
        }
      });
    }, 300);
  }

  /**
   * Add resize functionality to an element
   * @param element The element to make resizable
   * @param handle The resize handle element
   */
  private addResizeFunctionality(element: HTMLElement, handle: HTMLElement): void {
    let isResizing = false;
    let startX = 0;
    let startY = 0;
    let startWidth = 0;
    let startHeight = 0;
    let startLeft = 0;
    let startTop = 0;
    let aspectRatio = 1;
    
    const startResize = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      // Select the element when starting resize
      if (!element.classList.contains('selected')) {
        // Deselect any previously selected element
        if (this.selectedElement) {
          this.selectedElement.classList.remove('selected');
        }
        element.classList.add('selected');
        this.selectedElement = element;
      }
      
      isResizing = true;
      startX = e.clientX;
      startY = e.clientY;
      
      // Get starting dimensions from computed styles
      const computedStyle = window.getComputedStyle(element);
      startWidth = parseInt(computedStyle.width, 10);
      startHeight = parseInt(computedStyle.height, 10);
      startLeft = parseInt(computedStyle.left, 10) || 0;
      startTop = parseInt(computedStyle.top, 10) || 0;
      
      // Calculate aspect ratio
      aspectRatio = startWidth / startHeight;
      
      // Set absolute positioning if not already set
      if (computedStyle.position !== 'absolute') {
        const rect = element.getBoundingClientRect();
        const parentRect = element.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
        
        element.style.position = 'absolute';
        element.style.left = `${rect.left - parentRect.left}px`;
        element.style.top = `${rect.top - parentRect.top}px`;
        element.style.width = `${rect.width}px`;
        element.style.height = `${rect.height}px`;
      }
      
      // Add event listeners for resize
      document.addEventListener('mousemove', resize);
      document.addEventListener('mouseup', endResize);
      
      // Add a class to indicate resizing
      element.classList.add('resizing');
    };
    
    const resize = (e: MouseEvent) => {
      if (!isResizing) return;
      
      e.preventDefault();
      
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      
      // Handle rotation separately
      if (handle.classList.contains('rotate-handle')) {
        this.handleRotation(element, e);
        return;
      }
      
      // Handle resize based on which handle is being used
      if (handle.classList.contains('top-left')) {
        // Top-left corner - resize width and height, move position
        const newWidth = Math.max(30, startWidth - dx);
        const newHeight = Math.max(30, startHeight - dy);
        element.style.width = `${newWidth}px`;
        element.style.height = `${newHeight}px`;
        element.style.left = `${startLeft + (startWidth - newWidth)}px`;
        element.style.top = `${startTop + (startHeight - newHeight)}px`;
      } 
      else if (handle.classList.contains('top-right')) {
        // Top-right corner - resize width, adjust height, move top
        const newWidth = Math.max(30, startWidth + dx);
        const newHeight = Math.max(30, startHeight - dy);
        element.style.width = `${newWidth}px`;
        element.style.height = `${newHeight}px`;
        element.style.top = `${startTop + (startHeight - newHeight)}px`;
      } 
      else if (handle.classList.contains('bottom-left')) {
        // Bottom-left corner - resize width and height, move left
        const newWidth = Math.max(30, startWidth - dx);
        const newHeight = Math.max(30, startHeight + dy);
        element.style.width = `${newWidth}px`;
        element.style.height = `${newHeight}px`;
        element.style.left = `${startLeft + (startWidth - newWidth)}px`;
      } 
      else if (handle.classList.contains('bottom-right')) {
        // Bottom-right corner - resize width and height
        const newWidth = Math.max(30, startWidth + dx);
        const newHeight = Math.max(30, startHeight + dy);
        element.style.width = `${newWidth}px`;
        element.style.height = `${newHeight}px`;
      }
      else if (handle.classList.contains('top-center')) {
        // Top center - adjust height and top position
        const newHeight = Math.max(30, startHeight - dy);
        element.style.height = `${newHeight}px`;
        element.style.top = `${startTop + (startHeight - newHeight)}px`;
      }
      else if (handle.classList.contains('bottom-center')) {
        // Bottom center - adjust height only
        const newHeight = Math.max(30, startHeight + dy);
        element.style.height = `${newHeight}px`;
      }
      else if (handle.classList.contains('middle-left')) {
        // Middle left - adjust width and left position
        const newWidth = Math.max(30, startWidth - dx);
        element.style.width = `${newWidth}px`;
        element.style.left = `${startLeft + (startWidth - newWidth)}px`;
      }
      else if (handle.classList.contains('middle-right')) {
        // Middle right - adjust width only
        const newWidth = Math.max(30, startWidth + dx);
        element.style.width = `${newWidth}px`;
      }
      
      // Special handling for images to maintain aspect ratio
      const img = element.querySelector('img');
      if (img) {
        img.style.width = '100%';
        img.style.height = '100%';
      }
    };
    
    const endResize = (e: MouseEvent) => {
      if (!isResizing) return;
      
      e.preventDefault();
      
      isResizing = false;
      document.removeEventListener('mousemove', resize);
      document.removeEventListener('mouseup', endResize);
      
      // Remove the resizing class
      element.classList.remove('resizing');
    };
    
    // Add event listener to the handle
    handle.addEventListener('mousedown', startResize);
  }

  /**
   * Handle rotation of an element
   * @param element The element to rotate
   * @param e The mouse event
   */
  private handleRotation(element: HTMLElement, e: MouseEvent): void {
    // Get the center of the element
    const rect = element.getBoundingClientRect();
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    
    // Calculate the angle between the center and current mouse position
    const angle = Math.atan2(e.clientY - centerY, e.clientX - centerX) * (180 / Math.PI);
    
    // Apply the rotation to the element
    element.style.transform = `rotate(${angle + 90}deg)`;
  }
  
  /**
   * Set up mutation observer to detect when tables are added to the editor
   */
  private observeContentChanges(): void {
    const editorCanvases = this.editorCanvasList.toArray();
    
    editorCanvases.forEach(canvasRef => {
      const canvas = canvasRef.nativeElement;
      
      // Create a MutationObserver to watch for added tables
      const observer = new MutationObserver(mutations => {
        let shouldMakeResizable = false;
        let shouldMakeDraggable = false;
        
        mutations.forEach(mutation => {
          if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
            // Check if any tables were added
            mutation.addedNodes.forEach(node => {
              if ((node as Element).querySelector && (node as Element).querySelector('.editor-table')) {
                shouldMakeResizable = true;
                shouldMakeDraggable = true;
              }
            });
          }
        });
        
        if (shouldMakeResizable) {
          this.makeTablesResizable();
        }
        
        if (shouldMakeDraggable) {
          this.makeTablesDraggable();
        }
      });
      
      // Start observing the editor content for changes
      observer.observe(canvas, {
        childList: true,
        subtree: true
      });
    });
  }
  
  /**
   * Handle keyboard shortcuts
   * @param event The keyboard event
   */
  handleKeyDown(event: KeyboardEvent): void {
    // Save the current selection whenever the user is typing or navigating
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      this.lastSelection = selection.getRangeAt(0).cloneRange();
    }
    
    if (event.ctrlKey) {
      switch (event.key.toLowerCase()) {
        case 'b':
          event.preventDefault();
          this.formatDoc('bold');
          break;
        case 'i':
          event.preventDefault();
          this.formatDoc('italic');
          break;
        case 'u':
          event.preventDefault();
          this.formatDoc('underline');
          break;
        case '+':
        case '=':
          event.preventDefault();
          this.zoomIn();
          break;
        case '-':
          event.preventDefault();
          this.zoomOut();
          break;
      }
    }
  }
  
  /**
   * Format the document with the given command
   * @param command The command to execute (bold, italic, underline, etc.)
   */
  formatDoc(command: string): void {
    document.execCommand(command, false);
    this.focusCurrentEditorCanvas();
  }
  
  /**
   * Change the font family of the selected text
   * @param event The selection change event
   */
  formatFont(event: Event): void {
    const selectElement = event.target as HTMLSelectElement;
    const fontFamily = selectElement.value;
    
    if (fontFamily) {
      document.execCommand('fontName', false, fontFamily);
      this.focusCurrentEditorCanvas();
      selectElement.selectedIndex = 0;
    }
  }
  
  /**
   * Change the font size of the selected text
   * @param event The selection change event
   */
  formatFontSize(event: Event): void {
    const selectElement = event.target as HTMLSelectElement;
    const fontSize = selectElement.value;
    
    if (fontSize) {
      document.execCommand('fontSize', false, fontSize);
      this.focusCurrentEditorCanvas();
      selectElement.selectedIndex = 0;
    }
  }
  
  /**
   * Change the font color of the selected text
   * @param event The color picker change event
   */
  formatFontColor(event: Event): void {
    const colorInput = event.target as HTMLInputElement;
    const color = colorInput.value;
    
    if (color) {
      document.execCommand('foreColor', false, color);
      this.focusCurrentEditorCanvas();
    }
  }
  
  /**
   * Handle paste events
   * @param event The paste event
   */
  handlePaste(event: ClipboardEvent): void {
    // Custom paste handling can be added here if needed
  }
  
  /**
   * Add a new page to the editor
   */
  addPage(): void {
    const newPageIndex = this.pages.length;
    this.pages.push(newPageIndex);
    
    // Focus the new page after it's added to the DOM
    setTimeout(() => {
      const canvasElements = this.editorCanvasList.toArray();
      if (canvasElements[newPageIndex]) {
        canvasElements[newPageIndex].nativeElement.focus();
      }
    });
    
    // Initialize margins for the new page
    this.initializePageMargins();
  }
  
  /**
   * Delete a page from the editor
   * @param pageIndex The index of the page to delete
   */
  deletePage(pageIndex: number): void {
    // Prevent deleting the last page
    if (this.pages.length <= 1) {
      return;
    }
    
    // Get current content of all pages
    const canvasElements = this.editorCanvasList.toArray();
    const contentArray: string[] = [];
    
    canvasElements.forEach(canvas => {
      contentArray.push(canvas.nativeElement.innerHTML);
    });
    
    // Remove the page from the array
    this.pages.splice(pageIndex, 1);
    
    // Update page indices
    this.pages = this.pages.map((_, i) => i);
    
    // Wait for the DOM to update, then restore content and focus
    setTimeout(() => {
      const updatedCanvasElements = this.editorCanvasList.toArray();
      
      // Restore content for remaining pages
      updatedCanvasElements.forEach((canvas, i) => {
        // Skip the deleted page
        const sourceIndex = i >= pageIndex ? i + 1 : i;
        if (contentArray[sourceIndex]) {
          canvas.nativeElement.innerHTML = contentArray[sourceIndex];
        }
      });
      
      // Set focus to the appropriate page
      const focusIndex = Math.min(pageIndex, this.pages.length - 1);
      if (updatedCanvasElements[focusIndex]) {
        updatedCanvasElements[focusIndex].nativeElement.focus();
      }
      
      // Make sure tables are resizable after page deletion
      this.makeTablesResizable();
      this.makeTablesDraggable();
    });
  }
  
  /**
   * Increase the zoom level
   */
  zoomIn(): void {
    if (this.zoomLevel < this.maxZoom) {
      this.zoomLevel += this.zoomStep;
    }
  }
  
  /**
   * Decrease the zoom level
   */
  zoomOut(): void {
    if (this.zoomLevel > this.minZoom) {
      this.zoomLevel -= this.zoomStep;
    }
  }
  
  /**
   * Focus the current editor canvas
   */
  private focusCurrentEditorCanvas(): void {
    // Get the canvas that currently has focus or default to the first one
    const activeElement = document.activeElement;
    const canvasElements = this.editorCanvasList.toArray();
    
    // If a canvas already has focus, maintain it
    if (activeElement && activeElement.classList.contains('editor-canvas')) {
      (activeElement as HTMLElement).focus();
    } else if (canvasElements.length > 0) {
      // Otherwise focus the first canvas
      canvasElements[0].nativeElement.focus();
    }
  }
  
  /**
   * Save the current selection range
   */
  saveCurrentSelection(): void {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      this.lastSelection = selection.getRangeAt(0).cloneRange();
    }
  }
  
  /**
   * Open the export modal
   */
  openExportModal(): void {
    this.showExportModal = true;
  }
  
  /**
   * Close the export modal
   */
  closeExportModal(): void {
    this.showExportModal = false;
  }
  
  /**
   * Open the fields modal and save current selection
   */
  openFieldsModal(): void {
    // Save the current selection range
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      this.lastSelection = selection.getRangeAt(0).cloneRange();
    }
    
    this.showFieldsModal = true;
  }
  
  /**
   * Close the fields modal
   */
  closeFieldsModal(): void {
    this.showFieldsModal = false;
  }
  
  /**
   * Open the table modal and save current selection
   */
  openTableModal(): void {
    // Save the current selection range
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      this.lastSelection = selection.getRangeAt(0).cloneRange();
    }
    
    // Reset grid selection
    this.gridRows = 1;
    this.gridCols = 1;
    
    this.showTableModal = true;
  }
  
  /**
   * Close the table modal
   */
  closeTableModal(): void {
    this.showTableModal = false;
  }
  
  /**
   * Update the grid selection based on mouse position
   * @param rows Number of rows selected
   * @param cols Number of columns selected
   */
  updateGridSelection(rows: number, cols: number): void {
    this.gridRows = rows;
    this.gridCols = cols;
  }
  
  /**
   * Insert a table at the cursor position with the selected dimensions
   */
  insertTable(): void {
    // If we have a saved selection, restore it
    if (this.lastSelection) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(this.lastSelection);
        
        // Create HTML for table without the delete button
        let tableHtml = '<div class="editor-table-wrapper" contenteditable="false"><table class="editor-table" style="border-collapse: collapse; width: 100%; border: 2px solid #000; margin: 10px 0; table-layout: fixed;">';
        
        // Add rows and cells
        for (let i = 0; i < this.gridRows; i++) {
          tableHtml += '<tr>';
          for (let j = 0; j < this.gridCols; j++) {
            tableHtml += '<td style="border: 1px solid #000; padding: 8px; min-width: 50px; min-height: 24px; position: relative; overflow: visible;" contenteditable="true">&nbsp;</td>';
          }
          tableHtml += '</tr>';
        }
        
        tableHtml += '</table></div>';
        
        // Delete any selected text
        selection.deleteFromDocument();
        
        // Insert the table with line breaks before and after
        document.execCommand('insertHTML', false, '<br>' + tableHtml + '<br>');
        
        // Apply resizers after a small delay to ensure DOM is updated
        setTimeout(() => {
          this.makeTablesResizable();
          this.makeTablesDraggable();
        }, 100);
      }
    }
    
    // Focus back on the editor and close the modal
    this.focusCurrentEditorCanvas();
    this.closeTableModal();
  }
  
  /**
   * Make all tables in the editor draggable
   */
  private makeTablesDraggable(): void {
    setTimeout(() => {
      const tableWrappers = document.querySelectorAll('.editor-table-wrapper');
      
      tableWrappers.forEach(wrapper => {
        // Skip tables that have already been initialized for dragging
        if ((wrapper as HTMLElement).dataset['draggableInitialized'] === 'true') {
          return;
        }
        
        // Mark as initialized to avoid duplicate event listeners
        (wrapper as HTMLElement).dataset['draggableInitialized'] = 'true';
        
        // Set up dragging
        let isDragging = false;
        let startX = 0;
        let startY = 0;
        let startLeft = 0;
        let startTop = 0;
        
        // Ensure the wrapper is properly positioned
        const wrapperEl = wrapper as HTMLElement;
        if (window.getComputedStyle(wrapperEl).position !== 'absolute') {
          wrapperEl.style.position = 'relative';
          wrapperEl.style.display = 'inline-block';
          wrapperEl.style.width = 'auto';
          wrapperEl.style.cursor = 'move';
        }
        
        // Add event listener for drag start
        wrapperEl.addEventListener('mousedown', (e) => {
          // Don't start dragging if clicked on a table cell or resize handle
          if ((e.target as HTMLElement).tagName === 'TD' || 
              (e.target as HTMLElement).tagName === 'TH' ||
              (e.target as HTMLElement).classList.contains('editor-table-resizer')) {
            return;
          }
          
          e.preventDefault();
          e.stopPropagation();
          
          isDragging = true;
          startX = e.clientX;
          startY = e.clientY;
          
          // Get starting position
          const computedStyle = window.getComputedStyle(wrapperEl);
          startLeft = parseInt(computedStyle.left, 10) || 0;
          startTop = parseInt(computedStyle.top, 10) || 0;
          
          // Set absolute positioning if not already set
          if (computedStyle.position !== 'absolute') {
            const rect = wrapperEl.getBoundingClientRect();
            const parentRect = wrapperEl.parentElement?.getBoundingClientRect() || { left: 0, top: 0 };
            
            wrapperEl.style.position = 'absolute';
            wrapperEl.style.left = `${rect.left - parentRect.left}px`;
            wrapperEl.style.top = `${rect.top - parentRect.top}px`;
          }
          
          // Add event listeners for drag
          document.addEventListener('mousemove', drag);
          document.addEventListener('mouseup', endDrag);
          
          // Add a class to indicate dragging
          wrapperEl.classList.add('dragging');
        });
        
        const drag = (e: MouseEvent) => {
          if (!isDragging) return;
          
          e.preventDefault();
          
          const dx = e.clientX - startX;
          const dy = e.clientY - startY;
          
          wrapperEl.style.left = `${startLeft + dx}px`;
          wrapperEl.style.top = `${startTop + dy}px`;
        };
        
        const endDrag = (e: MouseEvent) => {
          if (!isDragging) return;
          
          e.preventDefault();
          
          isDragging = false;
          document.removeEventListener('mousemove', drag);
          document.removeEventListener('mouseup', endDrag);
          
          // Remove the dragging class
          wrapperEl.classList.remove('dragging');
        };
      });
    }, 200);
  }
  
  /**
   * Make all tables in the editor resizable
   */
  private makeTablesResizable(): void {
    setTimeout(() => {
      const allTables = document.querySelectorAll('.editor-canvas table.editor-table');
      
      allTables.forEach(table => {
        const cells = table.querySelectorAll('td');
        
        // Add table-level resizer for the entire table width
        if (!table.querySelector('.table-resizer')) {
          const tableResizer = document.createElement('div');
          tableResizer.className = 'table-resizer';
          tableResizer.style.position = 'absolute';
          tableResizer.style.top = '0';
          tableResizer.style.right = '-5px';
          tableResizer.style.width = '10px';
          tableResizer.style.height = '100%';
          tableResizer.style.cursor = 'col-resize';
          tableResizer.style.backgroundColor = 'transparent';
          tableResizer.style.zIndex = '20';
          
          // Make sure the table has position relative for the absolute positioned resizer
          (table as HTMLElement).style.position = 'relative';
          
          table.appendChild(tableResizer);
          
          // Add event listeners for entire table resizing
          this.addTableResizerEvents(tableResizer);
        }
        
        cells.forEach(cell => {
          // Remove existing resizers if any
          const existingResizers = cell.querySelectorAll('.editor-table-resizer');
          existingResizers.forEach(r => r.remove());
          
          // Add horizontal resizer (for column width)
          const hResizer = document.createElement('div');
          hResizer.className = 'editor-table-resizer h-resize';
          hResizer.style.position = 'absolute';
          hResizer.style.top = '0';
          hResizer.style.right = '0';
          hResizer.style.width = '5px';
          hResizer.style.height = '100%';
          hResizer.style.cursor = 'col-resize';
          hResizer.style.backgroundColor = 'transparent';
          hResizer.style.zIndex = '10';
          
          cell.appendChild(hResizer);
          
          // Add vertical resizer (for row height)
          const vResizer = document.createElement('div');
          vResizer.className = 'editor-table-resizer v-resize';
          vResizer.style.position = 'absolute';
          vResizer.style.bottom = '0';
          vResizer.style.left = '0';
          vResizer.style.width = '100%';
          vResizer.style.height = '5px';
          vResizer.style.cursor = 'row-resize';
          vResizer.style.backgroundColor = 'transparent';
          vResizer.style.zIndex = '10';
          
          cell.appendChild(vResizer);
          
          // Add event listeners for resizing
          this.addHorizontalResizerEvents(hResizer);
          this.addVerticalResizerEvents(vResizer);
        });
      });
    }, 200); // Increased delay to ensure DOM is fully rendered
  }

  /**
   * Add event listeners for resizing the entire table
   * @param resizer The table resizer element
   */
  private addTableResizerEvents(resizer: HTMLElement): void {
    let startX: number;
    let startWidth: number;
    const table = resizer.parentElement as HTMLElement;
    
    const mouseDownHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      startX = e.clientX;
      startWidth = table.offsetWidth;
      
      // Add event listeners
      document.addEventListener('mousemove', mouseMoveHandler);
      document.addEventListener('mouseup', mouseUpHandler);
      
      // Add active class to make resizer visible during resize
      resizer.classList.add('active');
    };
    
    const mouseMoveHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      const width = startWidth + (e.clientX - startX);
      if (width > 100) { // Minimum width
        // Set the width on the table itself
        table.style.width = `${width}px`;
      }
    };
    
    const mouseUpHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      // Remove event listeners
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
      
      // Remove active class
      resizer.classList.remove('active');
    };
    
    // Add event listener
    resizer.addEventListener('mousedown', mouseDownHandler);
  }

  /**
   * Add event listeners for horizontal table column resizing
   * @param resizer The horizontal resizer element
   */
  private addHorizontalResizerEvents(resizer: HTMLElement): void {
    let startX: number;
    let startWidth: number;
    const cell = resizer.parentElement as HTMLElement;
    
    const mouseDownHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      startX = e.clientX;
      startWidth = cell.offsetWidth;
      
      // Add event listeners
      document.addEventListener('mousemove', mouseMoveHandler);
      document.addEventListener('mouseup', mouseUpHandler);
      
      // Add active class to make resizer visible during resize
      resizer.classList.add('active');
    };
    
    const mouseMoveHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      const width = startWidth + (e.clientX - startX);
      if (width > 30) { // Minimum width
        cell.style.width = `${width}px`;
        
        // Force table layout to update
        const table = cell.closest('table');
        if (table) {
          table.style.tableLayout = 'fixed';
        }
      }
    };
    
    const mouseUpHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      // Remove event listeners
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
      
      // Remove active class
      resizer.classList.remove('active');
    };
    
    // Add event listener
    resizer.addEventListener('mousedown', mouseDownHandler);
  }
  
  /**
   * Add event listeners for vertical table row resizing
   * @param resizer The vertical resizer element
   */
  private addVerticalResizerEvents(resizer: HTMLElement): void {
    let startY: number;
    let startHeight: number;
    const cell = resizer.parentElement as HTMLElement;
    
    const mouseDownHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      startY = e.clientY;
      startHeight = cell.offsetHeight;
      
      // Add event listeners
      document.addEventListener('mousemove', mouseMoveHandler);
      document.addEventListener('mouseup', mouseUpHandler);
      
      // Add active class to make resizer visible during resize
      resizer.classList.add('active');
    };
    
    const mouseMoveHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      const height = startHeight + (e.clientY - startY);
      if (height > 24) { // Minimum height
        // Find the row and set height on all cells in the row
        const row = cell.parentElement;
        if (row) {
          const cells = row.querySelectorAll('td');
          cells.forEach(tdCell => {
            (tdCell as HTMLElement).style.height = `${height}px`;
          });
        }
      }
    };
    
    const mouseUpHandler = (e: MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      // Remove event listeners
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
      
      // Remove active class
      resizer.classList.remove('active');
    };
    
    // Add event listener
    resizer.addEventListener('mousedown', mouseDownHandler);
  }
  
  /**
   * Insert a field at the cursor position
   * @param fieldName The name of the field to insert
   */
  insertField(fieldName: string): void {
    // If we have a saved selection, restore it
    if (this.lastSelection) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(this.lastSelection);
        
        // Get the computed style of the selection to match its formatting
        const currentElement = selection.focusNode?.parentElement;
        
        // Create the field placeholder with double curly braces
        const fieldPlaceholder = document.createElement('span');
        fieldPlaceholder.textContent = `{{${fieldName}}}`;
        
        // Apply the current styling if possible
        if (currentElement) {
          fieldPlaceholder.style.fontFamily = window.getComputedStyle(currentElement).fontFamily;
          fieldPlaceholder.style.fontSize = window.getComputedStyle(currentElement).fontSize;
          fieldPlaceholder.style.fontWeight = window.getComputedStyle(currentElement).fontWeight;
          fieldPlaceholder.style.fontStyle = window.getComputedStyle(currentElement).fontStyle;
          fieldPlaceholder.style.color = window.getComputedStyle(currentElement).color;
        }
        
        // Delete any selected text and insert the field placeholder
        selection.deleteFromDocument();
        
        // Instead of inserting the element, insert the text to maintain editability
        const textNode = document.createTextNode(`{{${fieldName}}}`);
        selection.getRangeAt(0).insertNode(textNode);
        
        // Move the cursor after the inserted field
        const range = document.createRange();
        range.setStartAfter(textNode);
        range.setEndAfter(textNode);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
    
    // Focus back on the editor and close the modal
    this.focusCurrentEditorCanvas();
    this.closeFieldsModal();
  }
  
  /**
   * Export the document in the selected format
   * @param format The format to export as (pdf or html)
   */
  exportAs(format: string): void {
    // Get all editor canvases
    const canvasElements = this.editorCanvasList.toArray();
    const contentArray: string[] = [];
    
    // Get content from each page
    canvasElements.forEach(canvas => {
      contentArray.push(canvas.nativeElement.innerHTML);
    });
    
    if (format === 'pdf') {
      this.exportAsPDF(contentArray);
    } else if (format === 'html') {
      this.exportAsHTML(contentArray);
    }
    
    // Close the modal after export
    this.closeExportModal();
  }
  
  /**
   * Export the content as a PDF
   * @param contentArray Array of HTML content from each page
   */
  private exportAsPDF(contentArray: string[]): void {
    // Create a hidden iframe to prepare the content for PDF conversion
    const iframe = document.createElement('iframe');
    iframe.style.display = 'none';
    document.body.appendChild(iframe);
    
    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    
    if (doc) {
      doc.open();
      doc.write('<html><head><title>Export Document</title>');
      doc.write('<style>');
      doc.write('body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }');
      
      // Add border styling if enabled
      if (this.margins.showBorder) {
        doc.write(`.page { page-break-after: always; width: 210mm; min-height: 297mm; padding: ${this.margins.top}cm ${this.margins.right}cm ${this.margins.bottom}cm ${this.margins.left}cm; border: ${this.margins.borderWidth}px ${this.margins.borderStyle} ${this.margins.borderColor}; }`);
      } else {
        doc.write(`.page { page-break-after: always; width: 210mm; min-height: 297mm; padding: ${this.margins.top}cm ${this.margins.right}cm ${this.margins.bottom}cm ${this.margins.left}cm; }`);
      }
      
      doc.write('</style></head><body>');
      
      // Add each page with page breaks between them
      contentArray.forEach((content, index) => {
        doc.write(`<div class="page">${content}</div>`);
      });
      
      doc.write('</body></html>');
      doc.close();
      
      // In a real implementation, you would now convert the HTML to PDF
      // For this example, we'll just open the print dialog which allows saving as PDF
      setTimeout(() => {
        iframe.contentWindow?.print();
        // Clean up the iframe after printing
        setTimeout(() => {
          document.body.removeChild(iframe);
        }, 1000);
      }, 500);
    }
  }
  
  /**
   * Export the content as HTML
   * @param contentArray Array of HTML content from each page
   */
  private exportAsHTML(contentArray: string[]): void {
    // Create a complete HTML document
    let htmlContent = '<!DOCTYPE html>\n<html lang="en">\n<head>\n';
    htmlContent += '  <meta charset="UTF-8">\n';
    htmlContent += '  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n';
    htmlContent += '  <title>Exported Document</title>\n';
    htmlContent += '  <style>\n';
    htmlContent += '    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; display: flex; flex-direction: column; align-items: center; }\n';
    
    // Add border styling if enabled
    if (this.margins.showBorder) {
      htmlContent += `    .page { margin-bottom: 30px; width: 210mm; min-height: 297mm; padding: ${this.margins.top}cm ${this.margins.right}cm ${this.margins.bottom}cm ${this.margins.left}cm; border: 1px solid #ccc; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }\n`;
    }
    
    htmlContent += '  </style>\n';
    htmlContent += '</head>\n<body>\n';
    
    // Add each page
    contentArray.forEach((content, index) => {
      htmlContent += `  <div class="page">\n    ${content}\n  </div>\n`;
    });
    
    htmlContent += '</body>\n</html>';
    
    // Create a Blob with the HTML content
    const blob = new Blob([htmlContent], { type: 'text/html' });
    
    // Create a link element to download the HTML file
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'document.html';
    
    // Trigger the download
    document.body.appendChild(link);
    link.click();
    
    // Clean up
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  }
  
  /**
   * Open the margin settings modal
   */
  openMarginModal(): void {
    this.activeTab = 'margins-tab'; // Default to margins tab
    this.showMarginModal = true;
  }
  
  /**
   * Close the margin settings modal
   */
  closeMarginModal(): void {
    this.showMarginModal = false;
  }
  
  /**
   * Switch between tabs in the margin/border settings modal
   * @param tabId The ID of the tab to switch to
   */
  switchTab(tabId: string): void {
    this.activeTab = tabId;
  }
  
  /**
   * Apply margins and border to all editor canvases
   */
  applyMargins(): void {
    const canvasElements = this.editorCanvasList.toArray();
    
    canvasElements.forEach(canvas => {
      const element = canvas.nativeElement as HTMLElement;
      
      if (this.margins.showBorder) {
        // We need to create an inset border effect
        // First, remove any existing border
        element.style.border = 'none';
        
        // Clear any previous border container
        const existingBorder = element.querySelector('.border-container');
        if (existingBorder) {
          existingBorder.remove();
        }
        
        // Create a container for the border that's positioned according to margins
        const borderContainer = document.createElement('div');
        borderContainer.className = 'border-container';
        borderContainer.style.position = 'absolute';
        borderContainer.style.top = `${this.margins.top}cm`;
        borderContainer.style.bottom = `${this.margins.bottom}cm`;
        borderContainer.style.left = `${this.margins.left}cm`;
        borderContainer.style.right = `${this.margins.right}cm`;
        borderContainer.style.pointerEvents = 'none'; // To allow editing through the border
        borderContainer.style.zIndex = '1'; // Over content but under UI elements
        
        // Apply border to this container
        borderContainer.style.border = `${this.margins.borderWidth}px ${this.margins.borderStyle} ${this.margins.borderColor}`;
        
        // Make sure the editor canvas has a position for absolute positioning
        element.style.position = 'relative';
        
        // Apply padding to the content area to keep text away from the border
        element.style.paddingTop = `calc(${this.margins.top}cm + ${this.margins.borderWidth}px + 5px)`;
        element.style.paddingBottom = `calc(${this.margins.bottom}cm + ${this.margins.borderWidth}px + 5px)`;
        element.style.paddingLeft = `calc(${this.margins.left}cm + ${this.margins.borderWidth}px + 5px)`;
        element.style.paddingRight = `calc(${this.margins.right}cm + ${this.margins.borderWidth}px + 5px)`;
        
        // Add the border container to the editor canvas
        element.appendChild(borderContainer);
      } else {
        // If no border, just use regular margins
        element.style.paddingTop = `${this.margins.top}cm`;
        element.style.paddingBottom = `${this.margins.bottom}cm`;
        element.style.paddingLeft = `${this.margins.left}cm`;
        element.style.paddingRight = `${this.margins.right}cm`;
        element.style.border = '1px solid #ccc';
        
        // Remove any existing border container
        const existingBorder = element.querySelector('.border-container');
        if (existingBorder) {
          existingBorder.remove();
        }
      }
    });
    
    this.closeMarginModal();
  }
  
  onBorderChange(event: Event, side: 'Top' | 'Bottom' | 'Left' | 'Right', property: 'Width' | 'Style' | 'Color'): void {
    const propertyName = `border${side}${property}` as keyof Margins;
    
    if (property === 'Width') {
      const input = event.target as HTMLInputElement;
      const value = parseFloat(input.value);
      
      if (!isNaN(value) && value >= 0 && value <= 10) {
        (this.margins as any)[propertyName] = value;
      } else {
        input.value = (this.margins as any)[propertyName].toString();
      }
    } else if (property === 'Style') {
      const select = event.target as HTMLSelectElement;
      (this.margins as any)[propertyName] = select.value;
    } else if (property === 'Color') {
      const input = event.target as HTMLInputElement;
      (this.margins as any)[propertyName] = input.value;
    }
    
    this.updateMarginPreview();
  }

  /**
   * Handle common border property changes
   * @param event The input change event
   * @param property The border property (Width, Style, Color)
   */
  onCommonBorderChange(event: Event, property: 'Width' | 'Style' | 'Color'): void {
    const propertyName = `border${property}` as keyof Margins;
    
    if (property === 'Width') {
      const input = event.target as HTMLInputElement;
      const value = parseFloat(input.value);
      
      if (!isNaN(value) && value >= 0 && value <= 10) {
        this.margins.borderWidth = value;
        
        // Update all individual border widths to match
        this.margins.borderTopWidth = value;
        this.margins.borderBottomWidth = value;
        this.margins.borderLeftWidth = value;
        this.margins.borderRightWidth = value;
      } else {
        input.value = this.margins.borderWidth.toString();
      }
    } else if (property === 'Style') {
      const select = event.target as HTMLSelectElement;
      this.margins.borderStyle = select.value;
      
      // Update all individual border styles to match
      this.margins.borderTopStyle = select.value;
      this.margins.borderBottomStyle = select.value;
      this.margins.borderLeftStyle = select.value;
      this.margins.borderRightStyle = select.value;
    } else if (property === 'Color') {
      const input = event.target as HTMLInputElement;
      this.margins.borderColor = input.value;
      
      // Update all individual border colors to match
      this.margins.borderTopColor = input.value;
      this.margins.borderBottomColor = input.value;
      this.margins.borderLeftColor = input.value;
      this.margins.borderRightColor = input.value;
    }
    
    // Update the preview
    this.updateMarginPreview();
  }
  
  /**
   * Handle margin input changes
   * @param event The input change event
   * @param margin The margin property to update (top, bottom, left, right)
   */
  onMarginChange(event: Event, margin: 'top' | 'bottom' | 'left' | 'right'): void {
    const input = event.target as HTMLInputElement;
    const value = parseFloat(input.value);
    
    // Validate the input value
    if (!isNaN(value) && value >= 0 && value <= 10) {
      this.margins[margin] = value;
      
      // Update preview
      this.updateMarginPreview();
    } else {
      // Reset to previous valid value
      input.value = this.margins[margin].toString();
    }
  }
  
  /**
   * Handle border checkbox change
   * @param event The checkbox change event
   */
  onBorderToggle(event: Event): void {
    const checkbox = event.target as HTMLInputElement;
    this.margins.showBorder = checkbox.checked;
    this.updateMarginPreview();
  }
  
  /**
   * Handle border width change
   * @param event The input change event
   */
  onBorderWidthChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const value = parseFloat(input.value);
    
    // Validate border width (minimum 1, maximum 10)
    if (!isNaN(value) && value >= 1 && value <= 10) {
      this.margins.borderWidth = value;
      this.updateMarginPreview();
    } else {
      // Reset to previous valid value if input is invalid
      input.value = this.margins.borderWidth.toString();
    }
  }
  
  /**
   * Handle border style change
   * @param event The select change event
   */
  onBorderStyleChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    this.margins.borderStyle = select.value;
    this.updateMarginPreview();
  }
  
  /**
   * Handle border color change
   * @param event The color picker change event
   */
  onBorderColorChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.margins.borderColor = input.value;
    this.updateMarginPreview();
  }
  
  /**
   * Reset margins and border to default values
   */
  resetMargins(): void {
    // Reset to default values
    this.margins.top = this.defaultMargins.top;
    this.margins.bottom = this.defaultMargins.bottom;
    this.margins.left = this.defaultMargins.left;
    this.margins.right = this.defaultMargins.right;
    
    // Reset border settings
    this.margins.showBorder = this.defaultMargins.showBorder;
    this.margins.borderWidth = this.defaultMargins.borderWidth;
    this.margins.borderStyle = this.defaultMargins.borderStyle;
    this.margins.borderColor = this.defaultMargins.borderColor;

    this.margins.borderTopWidth = this.defaultMargins.borderTopWidth;
    this.margins.borderBottomWidth = this.defaultMargins.borderBottomWidth;
    this.margins.borderLeftWidth = this.defaultMargins.borderLeftWidth;
    this.margins.borderRightWidth = this.defaultMargins.borderRightWidth;
    
    this.margins.borderTopStyle = this.defaultMargins.borderTopStyle;
    this.margins.borderBottomStyle = this.defaultMargins.borderBottomStyle;
    this.margins.borderLeftStyle = this.defaultMargins.borderLeftStyle;
    this.margins.borderRightStyle = this.defaultMargins.borderRightStyle;
    
    this.margins.borderTopColor = this.defaultMargins.borderTopColor;
    this.margins.borderBottomColor = this.defaultMargins.borderBottomColor;
    this.margins.borderLeftColor = this.defaultMargins.borderLeftColor;
    this.margins.borderRightColor = this.defaultMargins.borderRightColor;
    
    // Update the margin preview
    this.updateMarginPreview();
  }
  
  /**
   * Update the margin and border preview
   */
  private updateMarginPreview(): void {
    // Update the content area border in the preview
    const contentAreas = document.querySelectorAll('.content-area');
    contentAreas.forEach(area => {
      const element = area as HTMLElement;
      if (this.margins.showBorder) {
        // Apply common border to all sides
        element.style.border = `${this.margins.borderWidth}px ${this.margins.borderStyle} ${this.margins.borderColor}`;
      } else {
        element.style.border = 'none';
      }
    });
  
    // Update margin indicators
    const marginIndicators = document.querySelectorAll('.margin-indicator');
    marginIndicators.forEach(indicator => {
      (indicator as HTMLElement).style.backgroundColor = this.margins.showBorder ? 'rgba(66, 133, 244, 0.3)' : 'transparent';
    });
  }
  
  /**
   * Initialize margins and border for newly added pages
   */
  private initializePageMargins(): void {
    // Wait for the DOM to update with the new page
    setTimeout(() => {
      const canvasElements = this.editorCanvasList.toArray();
      const lastCanvasElement = canvasElements[canvasElements.length - 1];
      
      if (lastCanvasElement) {
        const element = lastCanvasElement.nativeElement as HTMLElement;
        element.style.paddingTop = `${this.margins.top}cm`;
        element.style.paddingBottom = `${this.margins.bottom}cm`;
        element.style.paddingLeft = `${this.margins.left}cm`;
        element.style.paddingRight = `${this.margins.right}cm`;
        
        // Apply common border if enabled
        if (this.margins.showBorder) {
          element.style.border = `${this.margins.borderWidth}px ${this.margins.borderStyle} ${this.margins.borderColor}`;
        } else {
          element.style.border = '1px solid #ccc';
        }
      }
    });
  }
}