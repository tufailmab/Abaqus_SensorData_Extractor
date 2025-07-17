</head>
<body>
  <h1>Export Sensor Displacement Data (All Nodes)</h1>

  <p>
    This Python script extracts nodal displacement data (U1, U2, U3) from a selected part in an Abaqus <code>.odb</code> file and exports the results to organized text and Excel files. It automates sensor-based data extraction for extensometers, mesh sensitivity studies, and structural post-processing. (This tool was originally developed as part of a PhD research project in Pakistan. The affiliated university is intentionally not disclosed for confidentiality purposes).
  </p>

  <h2>Features</h2>
  <ul>
    <li><strong>Automatic ODB Detection:</strong> Opens the first <code>.odb</code> file in the working directory.</li>
    <li><strong>Multi-Node Data Export:</strong> Extracts displacement for each listed sensor node in U1, U2, and U3 directions.</li>
    <li><strong>Text and Excel Output:</strong> Saves each sensor's data in both <code>.txt</code> and <code>.xlsx</code> formats.</li>
    <li><strong>Excel Automation:</strong> Automatically formats and saves Excel files using Abaqus' Excel plugin.</li>
    <li><strong>Organized Output:</strong> Results are saved in a dedicated <code>Results</code> folder for easy access.</li>
  </ul>

  <h2>Usage Instructions</h2>
  <ol>
    <li>Edit the <code>part_name</code> and <code>node_numbers</code> list according to your model.</li>
    <li>Place the script in the same folder as your Abaqus <code>.odb</code> file.</li>
    <li>Run using Abaqus CAE:
      <pre><code>abaqus cae noGUI=ExportSensorData_All.py</code></pre>
    </li>
    <li>Find exported data in the <code>Results</code> folder.</li>
  </ol>

  <h2>Requirements</h2>
  <ul>
    <li>Abaqus CAE (for access to the Abaqus Python API)</li>
    <li>Python 2.7 (shipped with Abaqus)</li>
    <li><code>win32com</code> and <code>subprocess</code> (for Excel automation)</li>
    <li>A properly set ExcelUtilities plugin path</li>
  </ul>

  <h2>License</h2>
  <p>This project is licensed under the <strong>MIT License</strong>.</p>
  <p>You may use, modify, and distribute the script with attribution.</p>

  <h2>Developer Information</h2>
  <ul>
    <li><strong>Developer:</strong> Engr. Tufail Mabood</li>
    <li><strong>Contact:</strong> <a href="https://wa.me/+923440907874">WhatsApp</a></li>
    <li><strong>Note:</strong> Shared as an open-source tool to support Abaqus users. Contributions are welcome.</li>
  </ul>
</body>
</html>
