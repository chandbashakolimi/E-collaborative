'''''>>>>>>>>>> Tracker Data <<<<<<<<<<<<**

DataTable.Value("Title","Global")=glbTracker
DataTable.Value("ArtifactTitle","Global")=glbArtifact



'''>>>>>>>>>> Task Data <<<<<<<<<<<<**

DataTable.Value("TaskName","Global")=glbTask  
Dim Start_Date,Due_Date
Start_Date=(Month(date)&"/"&day(date)&"/"&year(date))
Due_Date=(Month(date)&"/"&day(date)&"/"&year(date))
DataTable.Value("StartDate","Global")=Start_Date
DataTable.Value("DueDate","Global")=Due_Date


''>>>>>>>>>> Source Code Data <<<<<<<<<<<<**

DataTable.Value("RepositoryName","Global")=glbRepo
DataTable.Value("DirectoryName","Global")=glbDirectory

'>>>>>>>>> Kanban Data <<<<<<<<<<<<**

DataTable.Value("KanbanName","Global")=glbkanban
DataTable.Value("Description","Global")=fGenerateRandomString(5)

'''>>>>>>>>> File Release Data <<<<<<<<<<<<**

DataTable.Value("PackageName","Global")=glbPackage
DataTable.Value("ReleaseName","Global")=glbRelease
'DataTable.Value("Title","Global")=glbTracker&"_REl"
'DataTable.Value("ArtifactTitle","Global")=glbArtifact
'

'>>>>>>>>> Document Data <<<<<<<<<<<<**
	DataTable.Value("DocPath","Global")="C:\Users\"&Environment.Value("UserName")&"\E collabe Automation\Sample Doc.txt"
	DataTable.Value("FolderName","Global")=glbFolder
	
	Dim url,SelEnv,glburl
	FSelectURL(EnvironmentFlag)	
	FwriteResultToHTML

'''################################################################################################################
''		Create new tracker
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create New Tracker]", oneIteration
	FeedDT "TRACKER_Create New Tracker"
	Datatable.Value("EndTime","Action1") = time()

''################################################################################################################
''		Create New Artifact
'################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create New Artifact]", oneIteration
	FeedDT "TRACKER_Create New Artifact"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''		Display the Tracker Summary
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Display the Trackers Summary page and display a tracker]", oneIteration
	FeedDT "TRACKER_Display the Trackers Summary page and display a tracker"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	View Graph of Tracker
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()

	RunAction "Action1 [View Graph of Tracker]", oneIteration
	FeedDT "TRACKER_View Graph of Tracker"
	Datatable.Value("EndTime","Action1") = time()
	
''''################################################################################################################
'	'	Import artifacts into a Tracker from a file
''''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Import artifacts into a Tracker from a file]", oneIteration
	FeedDT "TRACKER_Import artifacts into a Tracker from a file"
	Datatable.Value("EndTime","Action1") = time()
	
'	
'''	################################################################################################################
'''		Create New Artifact using Planning folder
'''	################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	
	RunAction "Action1 [Create New Artifact using Planning folder]", oneIteration
	FeedDT "TRACKER_Create New Artifact using Planning folder"
	Datatable.Value("EndTime","Action1") = time()

	

''################################################################################################################
	'Create New Transition
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	
	RunAction "Action1 [Create New Transition]", oneIteration
	FeedDT "TRACKER_Create New Transition"
	Datatable.Value("EndTime","Action1") = time()	


''################################################################################################################
	'Create Association
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	
	RunAction "Action1 [Check the Association and Create Association]", oneIteration
	FeedDT "TRACKER_Check the Association and Create Association"
	Datatable.Value("EndTime","Action1") = time()



''################################################################################################################
	'Create Parent Dependencies
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create artifact Dependencies with parent Artifact]", oneIteration
	FeedDT "TRACKER_Create artifact Dependencies with parent Artifact"
	Datatable.Value("EndTime","Action1") = time()
	
	
''################################################################################################################
'	'Create child Dependencies
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create artifact Dependencies for child Artifact]", oneIteration
	FeedDT "TRACKER_Create artifact Dependencies for child Artifact"
	Datatable.Value("EndTime","Action1") = time()

'
''''################################################################################################################
''''	'Create New Task
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create New Task]", oneIteration
	FeedDT "TASK_Create New Task"
	Datatable.Value("EndTime","Action1") = time()



'################################################################################################################
'	'Create and browse new repository
'################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	
	RunAction "Action1 [Create and browse new repository]", oneIteration
	FeedDT "SOURCE-CODE_Create and browse new repository"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	'Create a new Directory Name with a @ in the name
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create a new Directory Name with a @ in the name]", oneIteration
	FeedDT "SOURCE-CODE_Create a new Directory Name with a @ in the name"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	'Display Kanban List
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Display Kanban List]", oneIteration
	FeedDT "KANBAN_Display Kanban List"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	Create kanban and display the kanban List
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create kanban and display the kanban List]", oneIteration
	FeedDT "KANBAN_Create kanban and display the kanban List"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	Manage Mapping
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Manage Mapping]", oneIteration
	FeedDT "KANBAN_Manage Mapping"
	Datatable.Value("EndTime","Action1") = time()


''################################################################################################################
''	Group by priority
''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Group by priority]", oneIteration
	FeedDT "KANBAN_Group by priority"
	Datatable.Value("EndTime","Action1") = time()

'	
'''################################################################################################################
'''	create new release in new package
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [create new release in new package]", oneIteration
	FeedDT "FILE_RELEASE_create new release in new package"
	Datatable.Value("EndTime","Action1") = time()


'''################################################################################################################
'''	Check a migrated release that has artifacts in tabs Reported and Fixed Tracker Artifacts
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Check a migrated release that has artifacts in tabs Reported and Fixed Tracker Artifacts]", oneIteration
	FeedDT "FILE_RELEASE_Check a migrated release that has artifacts in tabs Reported and Fixed Tracker Artifacts"
	Datatable.Value("EndTime","Action1") = time()


''''################################################################################################################
'''	Create new folder
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create new folder]", oneIteration
	FeedDT "DOCUMENT_Create new folder"
	Datatable.Value("EndTime","Action1") = time()
		
'
'''''################################################################################################################
'''''Create Association in new Document
'''''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create Association in new Document]", oneIteration
	FeedDT "DOCUMENT_Create Association in new Document"
	Datatable.Value("EndTime","Action1") = time()



''''################################################################################################################
'''''View summary of planning folder
''''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	
	RunAction "Action1 [Display the PF Summary page and display a PF]", oneIteration
	FeedDT "PLANNING-FOLDER_Display the PF Summary page and display a PF"
	Datatable.Value("EndTime","Action1") = time()	


'''################################################################################################################
'''Total Efforts and points
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Planning folder summary of a specific Planning Folder]", oneIteration
	FeedDT "PLANNING-FOLDER_Planning folder summary of a specific Planning Folder"
	Datatable.Value("EndTime","Action1") = time()	
	

'''################################################################################################################
'''Change Effort
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Change Effort]", oneIteration
	FeedDT "PLANNING-FOLDER_Change Effort"
	Datatable.Value("EndTime","Action1") = time()	

'''################################################################################################################
'''Create New Planning Folder
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Create New Planning Folder]", oneIteration
	FeedDT "PLANNING-FOLDER_Create New Planning Folder"
	Datatable.Value("EndTime","Action1") = time()	


'''################################################################################################################
'''Delete a Package
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Delete a Package]", oneIteration
	FeedDT "FILE_RELEASE_Delete a Package"
	Datatable.Value("EndTime","Action1") = time()	
	

'''################################################################################################################
'''Delete Planning Folder
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Delete Planning Folder]", oneIteration
	FeedDT "PLANNING-FOLDER_Delete Planning Folder"
	Datatable.Value("EndTime","Action1") = time()	



'''################################################################################################################
'''Search Everything
'''################################################################################################################
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Search Everything]", oneIteration
	FeedDT "SEARCH_Search Everything"
	Datatable.Value("EndTime","Action1") = time()	

'''################################################################################################################
'''Search Everything this site
'''################################################################################################################
	Browser("E-collab").Back
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Search Everything this site]", oneIteration
	FeedDT "SEARCH_Search Everything this site"
	Datatable.Value("EndTime","Action1") = time()	

'''################################################################################################################
'''Search document
'''################################################################################################################
	Browser("E-collab").Back
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Search document]", oneIteration
	FeedDT "SEARCH_Search document"
	Datatable.Value("EndTime","Action1") = time()	

'''################################################################################################################
'''Search By Jump to ID
'''################################################################################################################
	Browser("E-collab").Back
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Search By Jump to ID]", oneIteration
	FeedDT "SEARCH_Search By Jump to ID"
	Datatable.Value("EndTime","Action1") = time()	

'''################################################################################################################
'''Search by People Search
'''################################################################################################################
	Browser("E-collab").Back
	Action1RowCount=Action1RowCount+1
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	fAdditionalTestInfo()
	
	RunAction "Action1 [Search by People Search]", oneIteration
	FeedDT "SEARCH_Search by People Search"
	Datatable.Value("EndTime","Action1") = time()	


Public Sub FeedDT(ByVal Scenario)
	datatable.GetSheet("Action1").SetCurrentRow(Action1RowCount)
	Datatable.Value("SubAction","Action1") = Scenario
	DataTable.Value("TestCaseNumber","Action1")=Action1RowCount

End Sub

Public Sub fAdditionalTestInfo()
	Datatable.value("Enviornment","Action1") = EnvironmentFlag
	Datatable.Value("StartTime","Action1") = time()
	Datatable.Value("Scenario","Action1") = ""
	DataTable.Value("BrowserAndVesrion","Action1")=Browser("E-collab").GetROProperty("version")
	
End Sub



'FwriteResultToHTML
'DataTable.Export (strResPath&"\"&Environment.Value("TestName")&".xlsx")
'DataTable.ExportSheet (strResPath&"\"&Environment.Value("TestName")&".xlsx"),"Global"
fWriteResultsToExcel(strResPath&"\"&Environment.Value("TestName")&"_" & strTimeStampTime & ".xlsx")





