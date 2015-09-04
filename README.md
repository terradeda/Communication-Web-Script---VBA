# KML-File-Generator---VBA

**Background**

This repository contains a collection of VBA scripts for generating a Google earth compatible communication web. 
This web shows the network communications between repeaters and collectors throughout an AMI network and can be used to 
visualize network efficiency.

An example of the resultant communications web can be seen below

![Example Communications Web](/images/logo.png)


**How To Run**

The process for generating this web is as follows:

  1) Run the SQL Scripts (Not Included in this Repository) which collect information on network performance.
  
      I wrote three seperate scripts, one which gathers information on the location and status of the collectors,
      one for the repeaters and one which gathers daily summaries of network communications between collectors and repeaters
      
  2) Import the VBA modules and userforms into an excel file. The VBA modules include:
  
      CollectorRepeaterAssociations - The GUI which acts as an interface and manages the import of data into a spreadsheet
      
      ExportKML - The module responsible for writing the KML file from the data gathered by the SQL Scripts
      
      PleaseWait - A Userform which displays a "Please Wait" while the KML file is being written
      
  3) Run the 'CollectorRepeaterAssociations' script, this launches the User Interface
  
  4) Select the 'Use New Data for Mapping' check box
  
  5) Browse and select the files containing the information gathered from the three SQL Scripts
  
  6) Click the 'Run' Button

