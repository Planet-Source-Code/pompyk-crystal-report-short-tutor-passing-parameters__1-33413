<div align="center">

## crystal report short tutor \(passing parameters\)


</div>

### Description

This is a nice crystal report tutorial...which guides u to use crystal report in Visual Basic..It's short...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[pompyk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pompyk.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pompyk-crystal-report-short-tutor-passing-parameters__1-33413/archive/master.zip)





### Source Code

<h2>CRYSTAL REPORT :</h2>
<br>
<h3>USE:</h3>
->Design the report.<br>
->Add groupings that parts the data as required.<br>
->Adding subtotals to your report as per groupings.<br>
<br>
<h3>HOW TO USE IT:</h3>
->Design the report layout in the crystal reports design view.<br>
->Do all sorts of codings in Visual Basic that runs the report.<br><br>
<h3>
NOTE:</h3> <font color="red">
Crystal report is available in various versions..version 5, 6,7, 8 etc. etc. The functionality of the report varies from version to version. Version 8 is a bit buggy so avoid using it and also some errors may come up while connecting it with oracle. Designing a report in crystal report environment is easy .... u just use the wizard and rest is easy. Save the report in .rpt format and rest is to be left to Visual basic (for attaching it)</font>
<br><br>
<h3>ACCESSING DESIGN ENVIRONMENT (crystal report) :</h3><font color="red">
(may not work in Version 8)
To access the design environment - simply click the Add - Ins | Report Designer menu option. or independently run crystal report from program .... (outside VB) and make a report using odbc,ado,etc etc. which is guided by the wizard (all sorts of report editing and grouping etc are available..all sorts of selection procedures are shown in the lists available) .... save the report in .rpt format [generally out of various options available the standard expert option is preferred]</font><br><h3>
USING THE ACTIVEX CRYSTAL REPORT CONTROL(from vb):</h3><br><b>
Open the Standard EXE window. Drag the Crystal Report Control ( from component select it first) and a command button on the form, size them and position them.. Open the property window of the Crystal Report and set the report file name by browsing and the selection formula as empty. Open the code window for the command button and enter the code.</b>
<br>
     <font color="green"> Private Sub Command1_Click()<br>
     CrystalReport1.Action = 1<br>
     End Sub<br>
<br>
.         ..and run it!!
<br>
<h3>
PASSING PARAMETERS IN CRYSTAL REPORT:</h3>
<font color="red">
yes, u can pass parameters in crystal report. For that u have to specify a condition from within visual basic. For example suppose user is the name of the table and age is the name of the field. Suppose we want to pass the age as a parameter then what we have to do is pass that value in the selectionformula property. (remember u use the crystal report activex and name it cr1 and also place a button so that parameter can be passed on clicking it and also u create a report1.rpt file using crytal report environment)
just see the below example: </font><br>
<br>
<font color="green"><br><br>
       Dim cond as String, a As Integer<br>
       ' below is the parameter which is passed..<br>
       a=inputbox("Enter your age plz..")
       'the below is the format for selection formula<br>
       cond="{user.age} = " & a & " "<br>
      With cr1<br>
         .Datafiles(0)= "C:\my documents\dpa.mdb"<br>
         .Destination = crptToWindow<br>
         .ReportFileName = "C:\my documents\report1.rpt"<br>
         .WindowState = crptMaximized<br>
      <br>    'assigning the selection formula
         .SelectionFormula = cond<br>
         .Action = 1<br>
       End With<br>
</font>
<br><h3>
WHY PEOPLE USE CRYSTAL REPORT:</h3><br><pre>
>limitation of data report since it can be only used when u implement ADO using data environment.
>Reports made using data reports are not graphically attractive.
>Data reports are quite tedious to implement.
VARIOUS TYPES OF CRYSTAL REPORTS: (OPTION)
The Listing Report:
  The listing report presents you with a series of four successive steps that guide you through the process of creating report. These steps are exactly the same as standard report the only difference is type of the report produced will just look like a list of information.
The Cross - tab Expert
  The cross tab expert presents you with a series of four successive steps. All except the third step, are the same as the standard expert. A cross tab is a spread sheet-style report that enables you to compare columns versus rows of data that you specify. This is a valuable report because you do not have to know how many data items are to appear on the x-axis before the report is run. One final thing to do is to select the summary field that will appear in the center data section on the report.
The Mail label expert
  The mail label expert presents you with a series of five successive steps in the process of creating labels. All steps, except the fourth step, are the same as the steps in the standard expert. The fourth step the label tab is used to select any one of the label styles given or you can define your own.
The Summary expert
  The summary expert presents you with a series of eight successive steps that guide you through the process of creating a summary report. All steps, except for the sixth step, are the same as the standard expert. The sixth step, the summary & Drilldown tab is used to define how to summarize and drill down on detail, based on groupings you've selected in the previous steps. Therefore you can choose whether or not to show these sections.
The Graph expert
  The graph expert presents you with a series of nine successive steps. Step 6 for the top n tab; is used to define how the report will select totals. You can select the top n number of records, where you specify the value for n. step 7, the graph tab, enables you to select the type of the graph to place on the report and to set attributes.
The Top N expert
  The Top N expert presents you with a series of eight successive steps that guide you through the process of creating hierarchical, detail-oriented report.
The Drill Down expert
  The Drill down expert presents you with a series of eight successive steps. It is essentially the same as Top N expert, except that you can choose which grouping sections to show or hide.</pre>
<h2>
author: Somdutt Ganguly</h2><br>
email: gangulysomdutt@yahoo.com<br>
date: 4/4/2002<br>

