# PCI-Calculator
A Python program that takes a CSV of road condition data and uses an Excel file to calculate the PCI values

<h3>Purpose</h3>
<p> This program will take a csv file of road condition samples and will import the data into an Excel file that calculates the Pavement Condition Index (PCI) value for each sample. It will then extract the calculated values from the Excel file. The PCI values are then appended to the end of the original data and a CSV file is exported to the original file's directory. </p>
<h3>Input File</h3>
<p> The Input file contains data collected by someone in the field during a road condition survey. It contains information such as the begin location, end location, sample length, sample width, and different types of distress.</p>
<p>The PCI value is calculated based on the Sample's Area and the amount of each type of distress ( Alligator Cracking, Transverse Cracking, Weathering, etc.). An example of an input file is shown below.</p>
<br>
<img src= "https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/Import%20File.PNG?raw=true" align="middle" height="342" width="653">

<h3>PCI Calculator Excel</h3>
<p> The Excel file containing the PCI calculator is complex. Each line in the Calc Many sheet shown below is a separate background sheet. Each row on the Calc Many sheet is a separate sample and contains the distress values for that sample.The highlighted cell in yellow is the calculated PCI value for that sample. </p>
<img src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/PCI%20Calculator%20Screenshot.PNG?raw=true" align="middle" height="342" width="653">
<br>
<p> There is a separate background sheet for each row on the Calc Many sheet. Below is an example of the background sheet for a sample. Due to the amount of calculations required on this sheet, I found it easier to import the values into the Excel file and extract the calculated value rather than trying to replicate the sheet in Python.</p>
<img src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/PCI%20Background%20Sheet.PNG?raw=true" align="middle" height="342" width="653">
<br>
<h3>Export File</h3>
<p>The exported file contains the information from the input file and appends the calculated PCI value to the end of each sample. The screenshot shown below is an example of the exported file and has the PCI value circled in red.</p>
<img src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/Export%20File.PNG?raw=true" align="middle" height="342" width="653">

<h3>Program</h3>
<p>The program is a simple GUI that will open a file browser when you click on the button. Once the file is selected, the GUI will add a continuous progress bar to show that it is responsive. The command window will show the progress to the user as shown below.</p>
<br>
<img  src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/Import1.PNG?raw=true" align="middle" height="342" width="653">
<img src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/Import2.PNG?raw=true" align="middle" height="342" width="653">
<br>
<p>Once the program is finished with calculating the values, the GUI will show the progress bar at 100% and will show the exported file location on the command window.</p>
<img src="https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/ImportFinished_LI.jpg?raw=true" align="middle" height="342" width="652">
<br>

<h3>Running the Program</h3>
<p>If you have python installed, you can run the program from the command line using <code>python3 PCICalculator.py</code></p>
<p>There is a <code>PCICalculator.exe</code> file in the dist folder. This exe was created on a Windows 10 system so that it could run on a Windows system without Python installed. <b>In either case, Excel will need to be installed on the computer.</b></p>
