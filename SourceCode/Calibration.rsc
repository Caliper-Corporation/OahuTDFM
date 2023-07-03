// Long Term Choice Models
Macro "Calibrate DriverLicense"(Args)
    opts = null
    opts.ModelName = "DriverLicense"
    opts.MacroName = "Driver License"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\DriverLicense.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate AutoOwnership"(Args)
    opts = null
    opts.ModelName = "AutoOwnership"
    opts.MacroName = "Auto Ownership"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\AutoOwnership.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkerCategory"(Args)
    opts = null
    opts.ModelName = "WorkerCategory"
    opts.MacroName = "Worker Category"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\WorkerCategory.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate DaycareStatus"(Args)
    opts = null
    opts.ModelName = "Daycare Participation"
    opts.MacroName = "Daycare Participation"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\DaycareStatus.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UniversityStatus"(Args)
    opts = null
    opts.ModelName = "Univ Participation"
    opts.MacroName = "Univ Participation"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UniversityStatus.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Frequency models
Macro "Calibrate WorkTourFreq"(Args)
    opts = null
    opts.ModelName = "Work Tours Frequency"
    opts.MacroName = "Work Tours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\WorkTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivTourFreq"(Args)
    opts = null
    opts.ModelName = "Univ Tours Frequency"
    opts.MacroName = "Univ Tours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UnivTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Duration Models
Macro "Calibrate FTWorkDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "FTWorkDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "FTWorkDuration1",
                        ModelFile: "FTWorkers_MandatoryDuration.mdl",
                        Filter: "WorkerCategory = 1 and WorkAttendance = 1", // Include WFH workers here
                        DestField: "WorkTAZ", 
                        Alternatives: Args.FTWorkDurAlts,
                        Utility: Args.FTWorkDurUtility,
                        ChoiceField: "Work1_DurChoice",
                        RandomSeed: 1499977}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\FullTimeWorkDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate PTWorkDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "PTWorkDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "PTWorkDuration1",
                        ModelFile: "PTWorkers_MandatoryDuration.mdl",
                        Filter: "WorkerCategory = 2 and WorkAttendance = 1",    // Note, WFH workers will have a duration
                        DestField: "WorkTAZ", 
                        Alternatives: Args.PTWorkDurAlts,
                        Utility: Args.PTWorkDurUtility,
                        ChoiceField: "Work1_DurChoice",
                        RandomSeed: 1899983}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\PartTimeWorkDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "UnivDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "UnivDuration1",
                        ModelFile: "Univ_MandatoryDuration.mdl",
                        Filter: "AttendUniv = 1",
                        DestField: "UnivTAZ", 
                        Alternatives: Args.UnivDurAlts,
                        Utility: Args.UnivDurUtility,
                        ChoiceField: "Univ1_DurChoice",
                        RandomSeed: 1699993}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UnivDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SchoolDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "SchoolDuration"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "SchoolDuration",
                        ModelFile: "School_Duration.mdl",
                        Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
                        DestField: "SchoolTAZ", 
                        Alternatives: Args.SchDurAlts,
                        Utility: Args.SchDurUtility,
                        ChoiceField: "School_DurChoice",
                        RandomSeed: 2099963}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\SchoolDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Start Time Models
Macro "Calibrate FTWorkStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "FTWorkStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "FTWorkStartTime1",
                        ModelTag: "Work1",
                        ModelFile: "FTWorkers_StartTime1.mdl",
                        Filter: "WorkerCategory = 1 and WorkAttendance = 1", // This includes WFH
                        DestField: "WorkTAZ",
                        Alternatives: Args.FTWorkStartAlts,
                        Utility: Args.FTWorkStartUtility,
                        ChoiceField: "Work1_StartInt",
                        RandomSeed: 2199979} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\FullTimeWorkStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate PTWorkStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "PTWorkStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "PTWorkStartTime1",
                        ModelTag: "Work1",
                        ModelFile: "PTWorkers_StartTime1.mdl",
                        Filter: "WorkerCategory = 2 and WorkAttendance = 1 and AttendSchool <> 1",
                        DestField: "WorkTAZ",
                        Alternatives: Args.PTWorkStartAlts,
                        Utility: Args.PTWorkStartUtility,
                        ChoiceField: "Work1_StartInt",
                        RandomSeed: 2699999}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\PartTimeWorkStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "UnivStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "UnivStartTime1",
                        ModelTag: "Univ1",
                        ModelFile: "Univ_StartTime1.mdl",
                        Filter: "AttendUniv = 1 and WorkAttendance <> 1",
                        DestField: "UnivTAZ",
                        Alternatives: Args.UnivStartAlts,
                        Utility: Args.UnivStartUtility,
                        ChoiceField: "Univ1_StartInt",
                        RandomSeed: 2399993}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UnivStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SchoolStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "SchoolStartTime"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "SchoolStartTime",
                        ModelTag: "School",
                        ModelFile: "School_StartTime.mdl",
                        Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
                        DestField: "SchoolTAZ", 
                        Utility: Args.SchStartUtility,
                        ChoiceField: "School_StartInt",
                        RandomSeed: 2899997}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\SchoolStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkStopsFreq"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Work Stops"
    opts.MacroName = "Mandatory Stops Choice"
    opts.MacroArgs = {Purpose: 'Work',
                        ToursView: objTours.GetView(),
                        Utility: Args.WorkStopsFreqUtility,
                        ChoicesField: "StopsChoice",
                        RandomSeed: 3599969,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\WorkStopsFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStopsFreq"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Univ Stops"
    opts.MacroName = "Mandatory Stops Choice"
    opts.MacroArgs = {Purpose: 'Univ',
                        ToursView: objTours.GetView(),
                        Utility: Args.UnivStopsFreqUtility,
                        ChoicesField: "StopsChoice",
                        RandomSeed: 3699961,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UnivStopsFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkStopsDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Stops_Work_Return"
    opts.MacroName = "Mandatory Duration Choice"
    opts.MacroArgs = {Type: 'Work',
                       Direction: 'Return',
                       ToursView: objTours.GetView(),
                       RandomSeed: 4099992,
                       LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\WorkStopsDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStopsDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Stops_Univ_Return"
    opts.MacroName = "Mandatory Duration Choice"
    opts.MacroArgs = {Type: 'Univ',
                       Direction: 'Return',
                       ToursView: objTours.GetView(),
                       RandomSeed: 4100102,
                       LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\UnivStopsDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourFrequency"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)

    // Sub Tour Models calibration
    opts = null
    opts.ModelName = "SubTour Choice"
    opts.MacroName = "Eval Sub Tour Choice"
    opts.MacroArgs = {ToursView: objTours.GetView(), LeaveDataOpen: 1}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\SubTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "SubTourDuration"
    opts.MacroName = "Subtour Activity Time"
    opts.MacroArgs = {ModelName: "SubTourDuration",
                        ModelFile: "SubTourDuration.mdl",
                        ToursView: objTours.GetView(),
                        Filter: "SubTour = 1",
                        OrigField: "Destination",
                        DestField: "SubTourTAZ",
                        Availabilities: Args.SubTourDurAvail,
                        Utility: Args.SubTourDurUtility,
                        ChoiceTable: Args.MandatoryTours,
                        ChoiceField: "SubTourActDurInt",
                        RandomSeed: 4399987,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\SubTourDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourStart"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)

    // Get Availability based on duration
    altSpec = Args.SubTourStartAlts
    availSpec = RunMacro("Get SubTour St Avails", altSpec)
    
    opts = null
    opts.ModelName = "SubTourStartTime"
    opts.MacroName = "Subtour Activity Time"
    opts.MacroArgs = {ModelName: "SubTourStartTime",
                        ModelFile: "SubTourStartTime.mdl",
                        ToursView: objTours.GetView(),
                        Filter: "SubTour = 1",
                        OrigField: "Destination",
                        DestField: "SubTourTAZ",
                        Availabilities: availSpec,
                        Alternatives: altSpec,
                        Utility: Args.SubTourStartUtility,
                        ChoiceTable: Args.MandatoryTours,
                        ChoiceField: "SubTourActStartInt",
                        RandomSeed: 4499969,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\SubTourStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate Model"(Args, Opts)
    abm = RunMacro("Get ABM Manager", Args)
    objT = CreateObject("Table", Args.AccessibilitiesOutputs)
    objDC = CreateObject("Table", Args.MandatoryDestAccessibility)
    
    modelName = Opts.ModelName
    calibrationFile = Opts.CalibrationFile
    if !GetFileInfo(calibrationFile) then
        Throw("Calibration file for " + modelName + " not found in the 'Data\\Calibration' folder.")

    // Run provided model first, that creates the model from the PME and opens all the relevant files.
    if Opts.MacroArgs <> null then
        RunMacro(Opts.MacroName, Args, Opts.MacroArgs)
    else
        RunMacro(Opts.MacroName, Args)

    // Retrieve Model Specification saved into Args array by the previous step
    modelSpec = Args.(modelName + " Spec")

    // Open Matrix Sources in model Spec
    for src in modelSpec.MatrixSources do
        mObjs.(src.Label) = CreateObject("Matrix", src.FileName)
    end

    // Call macro to Adjust ASC
    RunMacro("Calibrate ASCs", Opts, modelSpec)

    // Open calibration file in an editor
    shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)
    
    Return(1)
endMacro


Macro "Calibrate ASCs"(Opts, modelSpec)
    if modelSpec = null then
        Throw("Please run appropriate model step in the flowchart before running the calibration tool.")
        
    // Open calibration file and get alternatives, targets and thresholds
    objTable = CreateObject("AddTables", {TableName: Opts.CalibrationFile})
    vwC = objTable.TableView
    vecs = GetDataVectors(vwC + "|", {'Alternative', 'TargetShare'}, {OptArray: 1})

    alts = v2a(vecs.Alternative)    
    targets = v2a(vecs.TargetShare)
    thresholds = targets.Map(do (f) Return(0.02*f) end)
    max_iters = 20
    
    isAggregate = modelSpec.isAggregateModel
    
    // Copy input file into output
    inputModel = modelSpec.ModelFile
    pth = SplitPath(inputModel)
    outputModel = pth[1] + pth[2] + pth[3] + "_out" + pth[4]
    CopyFile(inputModel, outputModel)

    // Add output fields
    modify = CreateObject("CC.ModifyTableOperation", vwC)
    modify.FindOrAddField("InitialShare", "Real", 12, 2,)
    modify.FindOrAddField("InitialASC", "Real", 12, 4,)
    modify.FindOrAddField("AdjustedShare", "Real", 12, 2,)
    modify.FindOrAddField("AdjustedASC", "Real", 12, 4,)
    modify.Apply()
    
    dim shares[alts.length]
    convergence = 0
    iters = 0

    // Get Initial ASCs
    initialAscs = RunMacro("GetASCs", outputModel, alts)
    
    pbar = CreateObject("G30 Progress Bar", "Calibration Iterations...", true, max_iters)
    while convergence = 0 and iters <= max_iters do

        // Run Model
        outFile = RunMacro("Evaluate Model", modelSpec, outputModel)
        if outFile = null or !GetFileInfo(outFile) then
            Throw("Evaluating model for calibration failed")
            
        // Get Model Shares
        RunMacro("Generate Model Shares", alts, outFile, isAggregate, &shares)
        if iters = 0 then
            initialShares = CopyArray(shares)

        // Check convergence
        convergence = RunMacro("Convergence", shares, targets, thresholds)
        if convergence = null then
            Throw("Error in checking convergence")
            
        // Modify Model ASCs for next loop
        if convergence = 0 then
            RunMacro("Modify Model", outputModel, alts, shares, targets)

        iters = iters + 1

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Get Final ASCs
    finalAscs = RunMacro("GetASCs", outputModel, alts)

    vecsSet = null
    vecsSet.InitialShare = a2v(initialShares)
    vecsSet.AdjustedShare = a2v(shares)
    vecsSet.InitialASC = a2v(initialAscs)
    vecsSet.AdjustedASC = a2v(finalAscs)
    SetDataVectors(vwC + "|", vecsSet,)

    if convergence = 0 then
        ShowMessage("ASC Adjustment did not converge after " + i2s(max_iters) + " iterations")
    else
        ShowMessage("ASC Adjustment converged after " + i2s(iters) + " iterations")

endMacro


// Populates shares array with values from model run
Macro "Generate Model Shares"(alts, outFile, isAggregate, shares)
    if isAggregate then do
        m = OpenMatrix(outFile,)
        cores = GetMatrixCoreNames(m)
        stats = MatrixStatistics(m,)
        tsum = 0
        for i = 1 to stats.length do
            tsum = tsum + nz(stats.(cores[i]).Sum)
        end
        for i = 1 to alts.length do
            shares[i] = (100 * stats.(alts[i]).Sum)/tsum
        end
        m = null
    end
    else do // Disaggregate
        vw = OpenTable("Prob", "FFB", {outFile})
        probFlds = alts.Map(do (f) Return("[" + f + " Probability]") end)
        vecs = GetDataVectors(vw + "|", probFlds,)
        tsum = 0
        for i = 1 to vecs.length do // Note vecs array same length as alts
            shares[i] = VectorStatistic(vecs[i], "Sum",)
            tsum = tsum + shares[i]
        end
        for i = 1 to alts.length do
            shares[i] = (100 * shares[i])/tsum
        end
        CloseView(vw)
    end
endMacro


// ***** Generic macros used by all calibration macros *****
// Run Model. Uses the modelSpec array. Note that all relevant files are open at this point.
Macro "Evaluate Model"(modelSpec, outputModel)
    isAggregate = modelSpec.isAggregateModel

    o = CreateObject("Choice.Mode")
    o.ModelFile = outputModel
    
    if isAggregate then do
        probFile = GetRandFileName("Probability*.mtx")
        outputFile = GetRandFileName("Totals*.mtx")
        o.AddMatrixOutput("*", {Probability: probFile, Totals: outputFile})
    end
    else do
        outputFile = GetRandFileName("Probability*.bin")
        o.OutputProbabilityFile = outputFile
        o.OutputChoiceFile = GetRandFileName("Choices*.bin")
    end

    o.Run()
    Return(outputFile)
endMacro


// Returns 1 if all of the current model shares are within the target range.
Macro "Convergence"(shares, targets, thresholds)
    if shares.length <> targets.length or shares.length <> thresholds.length then
        Return()
        
    for i = 1 to shares.length do
        if shares[i] > (targets[i] + thresholds[i]) or shares[i] < (targets[i] - thresholds[i]) then // Out of bounds. Not Converged.
            Return(0)  
    end
    
    Return(1)    
endmacro


// Adjust ASC in output model
Macro "Modify Model"(model_file, alts, shares, targets)
    model = CreateObject("NLM.Model")
    model.Read(model_file, true)
    seg = model.GetSegment("*")

    for i = 1 to alts.length do
        alt = seg.GetAlternative(alts[i])
        if targets[i] > 0 then
            alt.ASC.Coeff = nz(alt.ASC.Coeff) + 0.5*log(targets[i]/shares[i])
        model.Write(model_file)
    end

    model.Clear()
endMacro


Macro "GetASCs"(model_file, alts)
    model = CreateObject("NLM.Model")
    model.Read(model_file, true)
    seg = model.GetSegment("*")

    dim ascs[alts.length]
    for i = 1 to alts.length do
        alt = seg.GetAlternative(alts[i])
        ascs[i] = alt.ASC.Coeff
    end

    model.Clear()
    Return(ascs)         
endMacro
