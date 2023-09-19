Macro "Pattern Choice Model"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    RunMacro("NM Models Setup", Args, abm)

    RunMacro("Sub Pattern Choice", Args, abm)

    Return(true)
endmacro


// Add several fields to HH and Person databases
Macro "NM Models Setup"(Args, abm)
    // Add several HH fields
    flds = {{Name: "HHAvgFreeTime", Type: "Real", Description: "Average free time (hours) per HH member"},
            {Name: "SubPattern", Type: "String", Width: 3, Description: "SubPattern:|N: No Tours|W: Work Tours|I: Individual Discretionary|J: Joint Discretionary"}}
    abm.DropHHFields({"HHAvgFreeTime"})
    abm.AddHHFields(flds) 
    
    // Compute Person and HH variables for pattern choice
    RunMacro("Calculate Time Usage Fields", Args, abm)

    // Create empty non-mandatory DC accessiblity table
    flds = {"Accessibility_JointOther", "Accessibility_JointShop", 
            "Accessibility_SoloOther", "Accessibility_SoloOther_NoLicense",
            "Accessibility_SoloShop", "Accessibility_SoloShop_NoLicense"}
    spec = {ReferenceFile: Args.TAZGeography, OutputFile: Args.NonMandatoryDestAccessibility, Fields: flds}
    objO = RunMacro("Create TAZ DC Logsum Table", spec)
endMacro


// Calculate variables needed by non mandatory choice models
Macro "Calculate Time Usage Fields"(Args, abm)
    abm.TimeManager = CreateObject("ABM.TimeManager", {TableName: Args.Persons, PersonID: abm.PersonID, HHID: abm.HHIDinPersonView})
    
    objT = CreateObject("Table", Args.MandatoryTours)
    vwTempTours = ExportView(objT.GetView() + "|", "MEM", "TempTours", {"PerID", "TourStartTime", "TourEndTime"},)
    abm.TimeManager.UpdateMatrixFromTours({ViewName: vwTempTours, PersonID: 'PerID', Departure: 'TourStartTime', Arrival: 'TourEndTime'})
    CloseView(vwTempTours)
    objT = null

    abm.TimeManager.ExportTimeUseMatrix(Args.MandTimeUseMatrix)
    
    // Fill HH Field with common and average time across persons
    hhSpec = {ViewName: abm.HHView, HHID: abm.HHID}
    opts = {HHSpec: hhSpec, Metric: "AverageFreeTime", HHFillField: "HHAvgFreeTime"}
    abm.TimeManager.FillHHTimeField(opts)
endMacro


Macro "Sub Pattern Choice"(Args, abm)
    // Run Model and populate results for worker on given day
    obj = CreateObject("PMEChoiceModel", {ModelName: "Sub Pattern Choice"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SubPatternChoice.mdl"
    obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
    obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "HH", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.SubPatternChoiceUtil, AvailabilityExpressions: Args.SubPatternChoiceAvail})
    obj.AddOutputSpec({ChoicesField: "SubPattern"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = 4699997
    ret = obj.Evaluate()
    Args.[SubPattern Spec] = CopyArray(ret)
    obj = null
endMacro
