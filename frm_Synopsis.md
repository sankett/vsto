# frm_Synopsis

The existing form  `frm_Synopsis`  is a Windows Forms application designed for a VSTO (Visual Studio Tools for Office) add-in. It interacts with Microsoft PowerPoint to create and manipulate slides. Below is a detailed explanation of its functionality:

#### Key Features:

1.  **Form Layout:**
    
    -   The form contains a  `GroupBox`  for entering the company name (`txtCompanyName`).
    -   Another  `GroupBox`  is used for entering a description (`txtDescription`).
    -   Buttons:
        -   `btnOk`: Inserts a new slide into the active PowerPoint presentation.
        -   `btnClose`: Closes the form.
        -   `btnClean`: Cleans up the text in the description field by removing extra spaces and line breaks.
2.  **PowerPoint Integration:**
    
    -   The form uses the  `Microsoft.Office.Interop.PowerPoint`  library to interact with PowerPoint.
    -   It opens a predefined PowerPoint template (`Synopsis2023.pptx`) to copy a chart from a specific slide.
    -   Inserts a new slide into the active PowerPoint presentation with a specific layout (`Title_Only`).
    -   Adds text and graphical elements (e.g., a chart and text boxes) to the new slide.
3.  **Text Processing:**
    
    -   The  `cleantext`  method removes extra spaces and line breaks from the description text.
    -   If the description text exceeds the height of a single text box, it splits the text into two text boxes.
4.  **Help Button:**
    
    -   Clicking the help button opens a URL in the default web browser for user guidance.
5.  **Error Handling:**
    
    -   The form includes basic error handling to display error messages in case of exceptions.

----------

### Detailed Requirements for Migration to Office.js

#### Functional Requirements:

1.  **Form UI:**
    
    -   Recreate the form UI using HTML, CSS, and JavaScript.
    -   Include input fields for the company name and description.
    -   Add buttons for "Insert," "Cancel," and "Clean."
2.  **PowerPoint Integration:**
    
    -   Use the Office.js API to interact with PowerPoint.
    -   Open a predefined PowerPoint template stored in a shared location.
    -   Copy a chart from a specific slide in the template.
    -   Insert a new slide into the active PowerPoint presentation.
    -   Add text and graphical elements to the new slide.
3.  **Text Processing:**
    
    -   Implement text cleaning functionality to remove extra spaces and line breaks.
    -   Split long descriptions into multiple text boxes if necessary.
4.  **Help Button:**
    
    -   Provide a link to the help guide that opens in a new browser tab.
5.  **Error Handling:**
    
    -   Implement robust error handling to display user-friendly error messages.

#### Non-Functional Requirements:

1.  **Cross-Platform Compatibility:**
    -   Ensure the add-in works on Windows, macOS, and Office Online.
2.  **Performance:**
    -   Optimize the add-in for quick response times.
3.  **Security:**
    -   Use secure methods to access and manipulate PowerPoint files.

----------

### Detailed Implementation Plan for Migration to Office.js

#### Step 1: Setup

-   Create a new Office Add-in project using the Yeoman generator for Office Add-ins.
-   Configure the project to support PowerPoint.

#### Step 2: UI Development

-   Design the form UI using HTML and CSS.
-   Use JavaScript to handle user interactions (e.g., button clicks).

#### Step 3: PowerPoint Integration

-   Use the Office.js API to:
    -   Open the predefined PowerPoint template.
    -   Copy a chart from a specific slide.
    -   Insert a new slide into the active presentation.
    -   Add text and graphical elements to the new slide.

#### Step 4: Text Processing

-   Implement JavaScript functions to clean and split text.

#### Step 5: Help Button

-   Add a hyperlink to the help guide.

#### Step 6: Error Handling

-   Use try-catch blocks to handle errors and display error messages.

#### Step 7: Testing

-   Test the add-in on Windows, macOS, and Office Online.
-   Validate functionality, performance, and compatibility.

#### Step 8: Deployment

-   Package the add-in and deploy it to the Office Add-ins store or a shared location.

----------

### Estimate Based on Tasks

#### Tasks:

-   **Setup:**
    -   Create Office Add-in project: 2 hours
    -   Configure project for PowerPoint: 1 hour
-   **UI Development:**
    -   Design form UI: 4 hours
    -   Implement user interactions: 3 hours
-   **PowerPoint Integration:**
    -   Open template and copy chart: 4 hours
    -   Insert new slide and add elements: 6 hours
-   **Text Processing:**
    -   Implement text cleaning: 2 hours
    -   Implement text splitting: 3 hours
-   **Help Button:**
    -   Add hyperlink: 1 hour
-   **Error Handling:**
    -   Implement error handling: 2 hours
-   **Testing:**
    -   Test on Windows, macOS, and Office Online: 6 hours
-   **Deployment:**
    -   Package and deploy add-in: 3 hours

#### Total Estimate:

-   **Total Hours:**  37 hours
-   **Buffer (20%):**  7.4 hours
-   **Grand Total:**  44.4 hours (~5.5 days)
