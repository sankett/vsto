---


---

<h1 id="frm_changechaptername">frm_ChangeChapterName</h1>
<p>The  <code>Frm_ChangeChapterName</code>  form is a Windows Form used in a VSTO (Visual Studio Tools for Office) add-in for PowerPoint. Its primary purpose is to allow users to rename a chapter or subchapter in a PowerPoint presentation. Below is a detailed breakdown of its functionality:</p>
<h4 id="key-features">Key Features:</h4>
<ol>
<li>
<p><strong>Form Initialization (<code>Frm_ChangeChapterName_Load</code>):</strong></p>
<ul>
<li>Retrieves the active PowerPoint presentation and the current slide index.</li>
<li>Determines whether the current slide is a “Chapter” or “Subchapter” based on the  <code>TAG_FLYPAGE_TYPE</code>  tag.</li>
<li>Displays the current name of the chapter or subchapter in the  <code>txtName</code>  text box by reading the  <code>TAG_FLYPAGE_NAME</code>  tag.</li>
</ul>
</li>
<li>
<p><strong>Help Button:</strong></p>
<ul>
<li>Clicking the help button opens a URL (<code>http://slv-hlweb/VAULT/PowerPointAddOnHelpGuide.aspx#chapter-options</code>) in the default web browser for user guidance.</li>
</ul>
</li>
<li>
<p><strong>Apply Button (<code>btnOk_Click</code>):</strong></p>
<ul>
<li>Updates the  <code>TAG_FLYPAGE_NAME</code>  tag of the current slide with the new name entered in the  <code>txtName</code>  text box.</li>
<li>Calls the  <code>InsertTOC3</code>  method to update the Table of Contents (TOC) based on the template properties.</li>
</ul>
</li>
<li>
<p><strong>Cancel Button (<code>btnClose_Click</code>):</strong></p>
<ul>
<li>Closes the form without making any changes.</li>
</ul>
</li>
<li>
<p><strong>Error Handling:</strong></p>
<ul>
<li>Displays error messages in a message box if exceptions occur during operations.</li>
</ul>
</li>
</ol>
<hr>
<h3 id="detailed-requirements-for-migration-to-office.js">Detailed Requirements for Migration to Office.js</h3>
<h4 id="functional-requirements">Functional Requirements:</h4>
<ol>
<li>
<p><strong>Initialization:</strong></p>
<ul>
<li>Retrieve the active PowerPoint presentation and the currently selected slide using Office.js APIs.</li>
<li>Determine if the slide is a “Chapter” or “Subchapter” based on custom tags.</li>
</ul>
</li>
<li>
<p><strong>UI Elements:</strong></p>
<ul>
<li>Create a web-based form with equivalent UI components:
<ul>
<li>A text box for entering the chapter/subchapter name.</li>
<li>Buttons for “Apply” and “Cancel.”</li>
<li>A help link to open the guidance URL.</li>
</ul>
</li>
</ul>
</li>
<li>
<p><strong>Apply Changes:</strong></p>
<ul>
<li>Update the custom tag (<code>TAG_FLYPAGE_NAME</code>) of the selected slide with the new name.</li>
<li>Update the Table of Contents dynamically.</li>
</ul>
</li>
<li>
<p><strong>Help Link:</strong></p>
<ul>
<li>Open the help URL in the default browser when the help link is clicked.</li>
</ul>
</li>
<li>
<p><strong>Error Handling:</strong></p>
<ul>
<li>Display error messages in a user-friendly manner (e.g., using modal dialogs).</li>
</ul>
</li>
</ol>
<hr>
<h3 id="detailed-implementation-plan-for-migration-to-office.js">Detailed Implementation Plan for Migration to Office.js</h3>
<h4 id="step-1-setup-office.js-environment">Step 1: Setup Office.js Environment</h4>
<ul>
<li>Create an Office Add-in project using the Yeoman generator for Office Add-ins.</li>
<li>Configure the manifest file to support PowerPoint.</li>
</ul>
<h4 id="step-2-design-the-web-based-form">Step 2: Design the Web-Based Form</h4>
<ul>
<li>Use HTML, CSS, and JavaScript to replicate the UI of the existing form.</li>
<li>Include:
<ul>
<li>A text input field for the chapter/subchapter name.</li>
<li>Buttons for “Apply” and “Cancel.”</li>
<li>A hyperlink for the help guide.</li>
</ul>
</li>
</ul>
<h4 id="step-3-implement-initialization-logic">Step 3: Implement Initialization Logic</h4>
<ul>
<li>Use the  <code>Office.context.document</code>  API to:
<ul>
<li>Retrieve the active presentation.</li>
<li>Get the currently selected slide.</li>
<li>Read custom tags (<code>TAG_FLYPAGE_TYPE</code>  and  <code>TAG_FLYPAGE_NAME</code>) to determine the slide type and populate the text box.</li>
</ul>
</li>
</ul>
<h4 id="step-4-implement-apply-button-logic">Step 4: Implement Apply Button Logic</h4>
<ul>
<li>Use the  <code>Office.context.document</code>  API to:
<ul>
<li>Update the  <code>TAG_FLYPAGE_NAME</code>  tag of the selected slide with the new name.</li>
<li>Call a function to update the Table of Contents dynamically.</li>
</ul>
</li>
</ul>
<h4 id="step-5-implement-help-link">Step 5: Implement Help Link</h4>
<ul>
<li>Add an event listener to the help link to open the guidance URL in a new browser tab.</li>
</ul>
<h4 id="step-6-implement-error-handling">Step 6: Implement Error Handling</h4>
<ul>
<li>Use try-catch blocks to handle errors and display them in modal dialogs.</li>
</ul>
<h4 id="step-7-testing-and-debugging">Step 7: Testing and Debugging</h4>
<ul>
<li>Test the add-in in PowerPoint to ensure all functionalities work as expected.</li>
<li>Debug any issues and refine the implementation.</li>
</ul>
<hr>
<h3 id="estimate-based-on-tasks">Estimate Based on Tasks</h3>
<h4 id="task-list">Task List:</h4>
<ul>
<li>
<p><strong>Setup Office.js Environment:</strong></p>
<ul>
<li>Install prerequisites (Node.js, Yeoman generator, etc.) - 2 hours</li>
<li>Create and configure the Office Add-in project - 2 hours</li>
</ul>
</li>
<li>
<p><strong>Design the Web-Based Form:</strong></p>
<ul>
<li>Create HTML structure - 3 hours</li>
<li>Style the form using CSS - 2 hours</li>
<li>Add JavaScript for UI interactions - 3 hours</li>
</ul>
</li>
<li>
<p><strong>Implement Initialization Logic:</strong></p>
<ul>
<li>Retrieve active presentation and slide - 2 hours</li>
<li>Read custom tags and populate the form - 3 hours</li>
</ul>
</li>
<li>
<p><strong>Implement Apply Button Logic:</strong></p>
<ul>
<li>Update custom tags - 3 hours</li>
<li>Update Table of Contents - 4 hours</li>
</ul>
</li>
<li>
<p><strong>Implement Help Link:</strong></p>
<ul>
<li>Add event listener for the help link - 1 hour</li>
</ul>
</li>
<li>
<p><strong>Implement Error Handling:</strong></p>
<ul>
<li>Add error handling for all operations - 2 hours</li>
</ul>
</li>
<li>
<p><strong>Testing and Debugging:</strong></p>
<ul>
<li>Test the add-in in PowerPoint - 4 hours</li>
<li>Debug and refine the implementation - 4 hours</li>
</ul>
</li>
</ul>
<h4 id="total-estimate">Total Estimate:</h4>
<ul>
<li><strong>Total Hours:</strong>  35 hours</li>
<li><strong>Total Days (8-hour workdays):</strong>  ~4.5 days</li>
</ul>

