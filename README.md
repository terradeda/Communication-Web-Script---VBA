# KML-File-Generator---VBA

**Background**

This repository contains a collection of VBA scripts for generating a Google Earth compatible communication web.I'm posting this here because I feel this script outlines a usefull method for plotting large amounts of data points and lines onto a map and can easily be modified to map various other types of data for various other applications. This script was written to help us visualize the network communications occuring between repeaters and collectors in our AMI network during any given day. 

Examples of the resultant communications web can be seen below

![Example Communications Web](https://cloud.githubusercontent.com/assets/11066939/9762504/ef21551c-56d1-11e5-86c0-213142a93efd.JPG)
![GUI for User Input](https://cloud.githubusercontent.com/assets/11066939/9762505/f2291736-56d1-11e5-9f20-c1891dfd4a39.JPG)
Legend: Green = Managed, Blue = Redundancy

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
      
  3) Run the 'CollectorRepeaterAssociations' script, this launches the user interface seen below...
  
  ![GUI for User Input](https://cloud.githubusercontent.com/assets/11066939/9762519/083258da-56d2-11e5-8d1c-4ca2da241f5b.JPG)
  
  4) Select the 'Use New Data for Mapping' check box
  
  5) Browse and select the files containing the information gathered from the three SQL Scripts
  
  6) Click the 'Run' Button
  
  7) This generates a 'Col-Rep Associations.kml' file which can be opened in a mapping program like google earth.

**How The KML File Is Generated**

A KML file is written in the Keyhole Markup Language. It is comprised of a list of predefined tags wrapped around text (Ex. < Desc>*Some Text*< /Desc>). The KML file generator uses these standard tags to write the bulk of the KML file and then simply injects the data pulled using the SQL scripts inside the corresponding tag. Because of this, modifications to this script can be easily made to add or alter functionality as the user sees fit.


