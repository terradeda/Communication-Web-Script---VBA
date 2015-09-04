# KML-File-Generator---VBA

**Background**

This repository contains a collection of VBA scripts for generating a Google earth compatible communication web. 
This web shows the network communications between repeaters and collectors throughout an AMI network and can be used to 
visualize network efficiency.

An example of the resultant communications web can be seen below

![Example Communications Web](/images/logo.png)


**How To Run**

The process for generating this web is as follows:

  1) Run the SQL Scripts (Not Included in this Repository) which collect information on network performance. There are three distinct scripts:
  
      One which gathers information on the location and status of the collectors
      
      One which gathers information on the location and status of the repeaters 
      
      One which gathers daily summaries of network communications between collectors and repeaters
      
  2) Import the VBA modules and userforms into an excel file. The VBA modules include:
  
      CollectorRepeaterAssociations - The GUI which acts as an interface and manages the import of data into a spreadsheet
      
      ExportKML - The module responsible for writing the KML file from the data gathered by the SQL Scripts
      
      PleaseWait - A Userform which displays a "Please Wait" while the KML file is being written
      
  3) Run the 'CollectorRepeaterAssociations' script, this launches the User Interface
  
  4) Select the 'Use New Data for Mapping' check box
  
  5) Browse and select the files containing the information gathered from the three SQL Scripts
  
  6) Click the 'Run' Button

**How The KML File Is Generated**

A KML file is written in the Keyhole Markup Language. It is comprised of a a list of predefined tags wrapped around text (Ex. <Desc>*Some Text*</Desc>). The KML file generator uses these standard tags to write the bulk of the KML file and then simply injects the data pulled using the SQL scripts inside the corresponding tag. Because of this, modifications to this script can be easily made to add or alter functionality as the user sees fit.
