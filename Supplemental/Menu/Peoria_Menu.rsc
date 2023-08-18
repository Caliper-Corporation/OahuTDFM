
// visualize menu items
Class "Visualize.Menu.Items"

    init do 
        self.runtimeObj = CreateObject("Model.Runtime")
    enditem 

    Macro "GetMenus" do
        Menus = {
              { ID: "PPMenu",  Title: "Persons Pivot" , Macro: "Menu_PP_Pivot" }
             ,{ ID: "HHMenu",  Title: "Household Pivot" , Macro: "Menu_HH_Pivot" }
             ,{ ID: "FlowMap", Title: "FlowMap" , Macro: "Menu_FlowMap" }
             ,{ ID: "PTFlows", Title: "PTFlowMap" , Macro: "Menu_PTFlowMap" }
             ,{ ID: "PTOn", Title: "Boarding Heatmap" , Macro: "Menu_PTOn" }
             ,{ ID: "PTOff", Title: "Alighting Heatmap" , Macro: "Menu_PTOff" }
             ,{ ID: "ConvChart", Title: "Convergence Chart" , Macro: "Menu_ConvergenceChart" }
             ,{ ID: "HeatMenu", Title: "Heatmap" , Macro: "Menu_HH_Heatmap" }
             ,{ ID: "ChordMenu", Title: "Chord Diagram" , Macro: "Menu_Chord_Diagram" }
            }
        
        Return(Menus)
    enditem 

    Macro "Menu_Chord_Diagram" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        paramName = self.runtimeObj.GetSelectedParamInfo().Name
        TAZGeoFile = self.runtimeObj.GetValue("WorkTAZ")
        self.runtimeObj.RunCode("ChordMap", tableArg, TAZGeoFile, paramName)
    enditem 

    Macro "Menu_HH_Pivot" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("HH_Pivot", tableArg)
    enditem 
   
    Macro "Menu_PP_Pivot" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("PP_Pivot", tableArg, opts)
    enditem 

    Macro "Menu_ConvergenceChart" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.ChartTitle = self.runtimeObj.GetSelectedParamInfo().Description
        self.runtimeObj.RunCode("ConvergenceChart", opts)
        enditem 

    macro "Menu_Chord" do 
        mName = self.runtimeObj.GetSelectedParamInfo().Value
        TAZGeoFile = self.runtimeObj.GetValue("TG_ZonalTable")
        self.runtimeObj.RunCode("CreateWebDiagram", {MatrixName: mName, TAZDB: TAZGeoFile, DiagramType: "Chord"})
    enditem         

    macro "Menu_Sankey" do 
        mName = self.runtimeObj.GetSelectedParamInfo().Value
        TAZGeoFile = self.runtimeObj.GetValue("TG_ZonalTable")
        self.runtimeObj.RunCode("CreateWebDiagram", {MatrixName: mName, TAZDB: TAZGeoFile, DiagramType: "Sankey"})
    enditem         

    macro "Menu_HH_Heatmap" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("CreateHeatmap", {TableName: tableArg})
    enditem

    Macro "Menu_FlowMap" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description
        opts.FlowFields = {"AB_Flow_PCE","BA_Flow_PCE"}
        opts.vocFields = {"AB_VOC","BA_VOC"}
        opts.LineLayer = self.runtimeObj.GetValue("WorkRoadDBD")
        opts.FlowsOnly = false
        opts.vocStyleFile = self.runtimeObj.GetValues().FlowMapStyles

        self.runtimeObj.RunCode("CreateFlowThemes", opts)
    enditem 
    
    Macro "Menu_PTFlowMap" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description
        opts.FlowFields = {"AB_TransitFlow","BA_TransitFlow"}
        opts.LineLayer = self.runtimeObj.GetValue("WorkRoadDBD")
        opts.FlowsOnly = true
        self.runtimeObj.RunCode("CreateFlowThemes", opts)
    enditem     

    // Boarding/Alighting heatmaps
    Macro "Menu_PTOn" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description + " - Boarding"
        opts.RS = self.runtimeObj.GetValue("WorkRS")        
        opts.Field = "On"
        self.runtimeObj.RunCode("OnOffHeatMap", opts)
    enditem     

    Macro "Menu_PTOff" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description + " - Alighting"
        opts.RS = self.runtimeObj.GetValue("WorkRS")        
        opts.Field = "Off"
        self.runtimeObj.RunCode("OnOffHeatMap", opts)
    enditem     
EndClass


// visualize menu items
// Main toolbar menues
MenuItem "TCRPC Menu Item" text: "TCRPC"
    menu "TCRPC Menu"
 
Menu "TCRPC Menu"
    init do
    enditem
 
    MenuItem "CreateScenario" text: "Create Scenario" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Create Scenario", Args)
    enditem
    MenuItem "Calibrate" text: "Calibrate Choice Models" menu "MENU_Calibration"
endMenu


menu "MENU_Calibration"
    MenuItem "Long Term Choices" text: "Long Term Choices" menu "LongTermChoices_Menu"
    Separator  
    MenuItem "Mandatory Tours" text: "Mandatory Tours" menu "MandatoryTours_Menu"  
    MenuItem "Mandatory Stops" text: "Mandatory Stops" menu "MandatoryStops_Menu"
    MenuItem "Mandatory SubTours" text: "Mandatory SubTours" menu "MandatorySubTours_Menu"
endmenu


menu "LongTermChoices_Menu"
    MenuItem "Driver License" text: "Driver License" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate DriverLicense", Args)
    endItem

    MenuItem "Auto Ownership" text: "Auto Ownership" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate AutoOwnership", Args)    
    endItem

    Separator

    MenuItem "Worker Category" text: "Worker Category" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkerCategory", Args)
    endItem

    Separator

    MenuItem "Daycare Status" text: "Daycare Status" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate DaycareStatus", Args)    
    endItem

    MenuItem "University Status" text: "University Status" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UniversityStatus", Args)    
    endItem
endmenu


menu "MandatoryTours_Menu"
    MenuItem "Work Tour Frequency" text: "Work Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkTourFreq", Args)    
    endItem

    MenuItem "Univ Tour Frequency" text: "University Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivTourFreq", Args)    
    endItem
    
    Separator

    MenuItem "FTWork1 Duration" text: "Full Time Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate FTWorkDuration", Args)   
    endItem

    MenuItem "PTWork1 Duration" text: "Part Time Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate PTWorkDuration", Args)   
    endItem

    MenuItem "Univ Duration" text: "University Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivDuration", Args)   
    endItem

    MenuItem "School Duration" text: "School Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SchoolDuration", Args)   
    endItem

    Separator

    MenuItem "FTWork1 Start" text: "Full Time Work Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate FTWorkStart", Args)   
    endItem

    MenuItem "PTWork1 Start" text: "Part Time Work Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate PTWorkStart", Args)   
    endItem

    MenuItem "Univ Start" text: "University Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStart", Args)   
    endItem

    MenuItem "School Start" text: "School Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SchoolStart", Args)   
    endItem

    Separator

    MenuItem "Work_Mode" text: "Work Tour Mode" disabled do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.WorkMC_Calibration = 1
        Args.WorkMC_Calibration = null
    endItem

    MenuItem "Univ_Mode" text: "University Tour Mode" disabled do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.UnivMC_Calibration = 1
        Args.UnivMC_Calibration = null
    endItem

    MenuItem "School_ModeF" text: "School Forward Mode" disabled do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.SchoolMCF_Calibration = 1
        Args.SchoolMCF_Calibration = null
    endItem

    MenuItem "School_ModeR" text: "School Return Mode" disabled do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.SchoolMCR_Calibration = 1
        Args.SchoolMCR_Calibration = null
    endItem
endmenu


// Mandatory Stops Menu
// Mandatory Stops Frequency
menu "MandatoryStops_Menu"
    MenuItem "WorkStopsFreq" text: "Stop Frequency: Work Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkStopsFreq", Args)  
    endItem

    MenuItem "UnivStopsFreq" text: "Stop Frequency: University Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStopsFreq", Args)  
    endItem

    Separator

// Mandatory Stops Duration
    MenuItem "WorkStopsDur" text: "Stop Duration: Work Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkStopsDuration", Args)  
    endItem

    MenuItem "UnivStopsDur" text: "Stop Duration: University Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStopsDuration", Args)  
    endItem
endmenu


menu "MandatorySubTours_Menu"
    MenuItem "SubTour_Freq" text: "Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourFrequency", Args)  
    endItem

    MenuItem "SubTour_Dur" text: "Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourDuration", Args)  
    endItem

    MenuItem "SubTour_Start" text: "Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourStart", Args)  
    endItem

    MenuItem "SubTour_Mode" text: "Mode" disabled do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
    endItem
endmenu
