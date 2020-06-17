# PCI-Calculator
A Python program that takes a CSV of road condition data and uses an Excel file to calculate the PCI values

<h3>Purpose</h3>
<p> This program will take a csv file of road condition samples and will import the data into an Excel file that calculates the Pavement Condition Index (PCI) value for each sample. It will then extract the calculated values from the Excel file. The PCI values are then appended to the end of the original data and a CSV file is exported to the original file's directory. </p>
<br>
<h3>Input File</h3>
<p> The Input file contains data collected by someone in the field during a road condition survey. It contains information such as the begin location, end location, sample length, sample width, and different types of distress.</p>
<p>The PCI value is calculated based on the Sample's Area and the amount of each type of distress ( Alligator Cracking, Transverse Cracking, Weathering, etc.). An example of an input file is shown below.</p>
<br>
<img src= "https://github.com/mbaker92/PCI-Calculator/blob/master/PCICalculator/screenshots/Import%20File.PNG?raw=true" align="middle" height="342" width="653">

<h3>PCI Calculator Excel</h3>

