---


---

<h1 id="frm_watermark">frm_Watermark</h1>
<p>The  <code>frm_Watermark</code>  form is a Windows Forms-based UI for a VSTO (Visual Studio Tools for Office) PowerPoint Add-in. It allows users to insert watermarks into PowerPoint slides. Below is a detailed explanation of its functionality:</p>
<ol>
<li>
<p><strong>Form Components</strong>:</p>
<ul>
<li><strong>ComboBox (<code>cboWatermark</code>)</strong>: Allows users to select predefined watermark text or choose “CUSTOM WATERMARK” to input their own text.</li>
<li><strong>Custom TextBox (<code>CustomWatermarkText</code>)</strong>: Visible only when “CUSTOM WATERMARK” is selected. Users can input custom watermark text here.</li>
<li><strong>Buttons (<code>btnOk</code>  and  <code>btnClose</code>)</strong>:
<ul>
<li><code>btnOk</code>: Inserts the selected or custom watermark into the PowerPoint slide.</li>
<li><code>btnClose</code>: Closes the form without making changes.</li>
</ul>
</li>
<li><strong>Help Button</strong>: Opens a help guide in the browser.</li>
</ul>
</li>
<li>
<p><strong>Key Functionalities</strong>:</p>
<ul>
<li><strong>Watermark Selection</strong>:
<ul>
<li>Predefined watermarks are populated in the ComboBox (<code>cboWatermark</code>) during form load.</li>
<li>If “CUSTOM WATERMARK” is selected, the  <code>CustomWatermarkText</code>  TextBox becomes visible for user input.</li>
</ul>
</li>
<li><strong>Watermark Insertion</strong>:
<ul>
<li>On clicking  <code>btnOk</code>, the selected or custom watermark is inserted into the PowerPoint slide using the  <code>Watermark.InsertWatermark</code>  method.</li>
<li>Validation ensures that a watermark is selected or entered before proceeding.</li>
</ul>
</li>
<li><strong>Help Guide</strong>:
<ul>
<li>Clicking the help button opens a URL in the browser for user assistance.</li>
</ul>
</li>
</ul>
</li>
<li>
<p><strong>Error Handling</strong>:</p>
<ul>
<li>Basic error handling is implemented in the  <code>btnOk_Click</code>  method, but exceptions are not logged or displayed to the user.</li>
</ul>
</li>
<li>
<p><strong>Logging</strong>:</p>
<ul>
<li>Log points (<code>AddLogPoint</code>) are added to track user actions like inserting a watermark.</li>
</ul>
</li>
</ol>
<hr>
<h3 id="detailed-requirements-for-migration-to-office.js">Detailed Requirements for Migration to Office.js</h3>
<ol>
<li>
<p><strong>UI Requirements</strong>:</p>
<ul>
<li>Replace the Windows Forms UI with an Office.js Task Pane Add-in.</li>
<li>Use HTML, CSS, and JavaScript to replicate the form’s layout and functionality.</li>
<li>Include a dropdown for predefined watermarks and a text input for custom watermarks.</li>
<li>Add buttons for “Insert” and “Cancel” actions.</li>
<li>Provide a link to the help guide.</li>
</ul>
</li>
<li>
<p><strong>Functional Requirements</strong>:</p>
<ul>
<li>Populate the dropdown with predefined watermarks.</li>
<li>Show/hide the custom watermark input field based on the dropdown selection.</li>
<li>Validate user input before inserting the watermark.</li>
<li>Insert the selected or custom watermark into the active PowerPoint slide.</li>
<li>Log user actions for analytics.</li>
</ul>
</li>
<li>
<p><strong>Technical Requirements</strong>:</p>
<ul>
<li>Use the Office.js API to interact with PowerPoint slides.</li>
<li>Implement the watermark insertion logic using PowerPoint’s slide manipulation capabilities.</li>
<li>Ensure compatibility with modern browsers and Office versions.</li>
</ul>
</li>
<li>
<p><strong>Error Handling</strong>:</p>
<ul>
<li>Display user-friendly error messages for validation and API errors.</li>
<li>Log errors for debugging purposes.</li>
</ul>
</li>
<li>
<p><strong>Help Guide</strong>:</p>
<ul>
<li>Open the help guide URL in a new browser tab.</li>
</ul>
</li>
</ol>
<hr>
<h3 id="detailed-implementation-plan-for-migration-to-office.js">Detailed Implementation Plan for Migration to Office.js</h3>
<h4 id="setup-office.js-add-in-project">1.  <strong>Setup Office.js Add-in Project</strong></h4>
<ul>
<li>Create a new Office Add-in project using the Yeoman generator or Visual Studio.</li>
<li>Configure the manifest file to support PowerPoint.</li>
</ul>
<h4 id="ui-development">2.  <strong>UI Development</strong></h4>
<ul>
<li>Design the Task Pane UI using HTML and CSS.</li>
<li>Implement the dropdown for predefined watermarks and a text input for custom watermarks.</li>
<li>Add “Insert” and “Cancel” buttons.</li>
<li>Include a link to the help guide.</li>
</ul>
<h4 id="populate-dropdown">3.  <strong>Populate Dropdown</strong></h4>
<ul>
<li>Use JavaScript to populate the dropdown with predefined watermarks.</li>
</ul>
<h4 id="event-handling">4.  <strong>Event Handling</strong></h4>
<ul>
<li>Add event listeners for dropdown selection changes to toggle the visibility of the custom watermark input field.</li>
<li>Add click event handlers for the “Insert” and “Cancel” buttons.</li>
</ul>
<h4 id="watermark-insertion-logic">5.  <strong>Watermark Insertion Logic</strong></h4>
<ul>
<li>Use the Office.js API to insert the selected or custom watermark into the active PowerPoint slide.</li>
<li>Implement validation to ensure a watermark is selected or entered.</li>
</ul>
<h4 id="logging">6.  <strong>Logging</strong></h4>
<ul>
<li>Implement logging for user actions and errors.</li>
</ul>
<h4 id="help-guide">7.  <strong>Help Guide</strong></h4>
<ul>
<li>Add functionality to open the help guide URL in a new browser tab.</li>
</ul>
<h4 id="testing">8.  <strong>Testing</strong></h4>
<ul>
<li>Test the add-in in different browsers and Office versions.</li>
<li>Validate the watermark insertion functionality.</li>
</ul>
<h4 id="deployment">9.  <strong>Deployment</strong></h4>
<ul>
<li>Package the add-in and deploy it to a test environment.</li>
<li>Update the manifest file for production deployment.</li>
</ul>
<hr>
<h3 id="estimate-and-task-breakdown">Estimate and Task Breakdown</h3>
<h4 id="tasks"><strong>Tasks</strong></h4>
<ul>
<li>
<p><strong>Setup and Configuration</strong>:</p>
<ul>
<li>Create Office.js project: 2 hours</li>
<li>Configure manifest file: 1 hour</li>
</ul>
</li>
<li>
<p><strong>UI Development</strong>:</p>
<ul>
<li>Design Task Pane layout: 4 hours</li>
<li>Implement dropdown and text input: 2 hours</li>
<li>Add buttons and help link: 2 hours</li>
</ul>
</li>
<li>
<p><strong>Functionality Implementation</strong>:</p>
<ul>
<li>Populate dropdown: 1 hour</li>
<li>Handle dropdown selection changes: 1 hour</li>
<li>Implement watermark insertion logic: 4 hours</li>
<li>Add validation: 2 hours</li>
</ul>
</li>
<li>
<p><strong>Logging and Error Handling</strong>:</p>
<ul>
<li>Implement logging: 2 hours</li>
<li>Add error handling: 2 hours</li>
</ul>
</li>
<li>
<p><strong>Testing</strong>:</p>
<ul>
<li>Functional testing: 4 hours</li>
<li>Cross-browser testing: 2 hours</li>
</ul>
</li>
<li>
<p><strong>Deployment</strong>:</p>
<ul>
<li>Package add-in: 1 hour</li>
<li>Deploy to test environment: 2 hours</li>
<li>Update manifest for production: 1 hour</li>
</ul>
</li>
</ul>
<h4 id="total-estimate-32-hours-4-working-days"><strong>Total Estimate</strong>: 32 hours (4 working days)</h4>

