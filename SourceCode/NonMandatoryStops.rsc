//==========================================================================================================
Macro "JointStops Setup"(Args)
    RunMacro("Discretionary Stops Setup", Args, {Type: 'Joint'})
    Return(true)
endMacro


Macro "JointStops Frequency"(Args)
    RunMacro("Discretionary Stops Freq", Args, {Type: 'Joint'})
    Return(true)
endMacro


Macro "JointStops Destination"(Args)
    RunMacro("Discretionary Stops Dest", Args, {Type: 'Joint'})
    Return(true)
endMacro


Macro "JointStops Duration"(Args)
    RunMacro("Discretionary Stops Dur", Args, {Type: 'Joint'})
    Return(true)
endMacro


//==========================================================================================================
Macro "JointStops Scheduling"(Args)
    maxTours = 3 // A HH can at most make three joint tours (O2S1) on a weekday
    abm = RunMacro("Get ABM Manager", Args)

    // Open Tours file and export to InMemory View
    obj = CreateObject("Table", Args.NonMandatoryJointTours)
    vw = obj.GetView()
    {flds, specs} = GetFields(vw,)
    vwMem = ExportView(vw + "|", "MEM", "JointToursFile",,)
    obj = null

    objT = CreateObject("Table", vwMem)
    // Scheduling prep by assigning tour number order to joint tours within the HH
    RunMacro("Tour Order", {View: vwMem, Type: 'Joint'})

    // Load time manager object
    TimeManager = RunMacro("Get Time Manager", abm)
    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.JointTimeUseMatrix}) // loading the existing file

    opt = null
    opt.abmManager = abm
    opt.ToursObj = objT
    opt.TimeManager = TimeManager
    pbar = CreateObject("G30 Progress Bar", "Running Intermediate Stops Scheduling for Joint Discretionary Tours (Max 3 tours)", false, maxTours)
    for i = 1 to maxTours do
        opt.TourNumber = i

        // Select records from the tour file
        filter = "TourOrder = " + string(i)
        objT.SelectByQuery({Query: filter, SetName: "__SelectedTours" + String(i)})
        opt.TourSet = "__SelectedTours" + String(i)
        
        pSet = RunMacro("Select Persons on Tour", opt)
        opt.PersonSet = pSet

        // *********** Determine arrival time from previous tour and departure time for next tour
        RunMacro("JointStop Scheduling Prep", opt)

        // *********** Run Forward Stop Scheduling
        RunMacro("Stop Scheduling Forward", Args, opt)

        // *********** Run Return Stop Scheduling
        RunMacro("Stop Scheduling Return", Args, opt)

        // *********** Update time manager
        RunMacro("JointStop Update TimeManager", opt)

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()
    obj = null

    // Export the updated in-memory view back to tour table
    ExportView(vwMem + "|", "FFB", Args.NonMandatoryJointTours, flds,)
    //ExportView(vwMem + "|", "FFB", Args.NonMandatoryJointTours, ,)
    CloseView(vwMem)

    // Export Final Time Manager Matrix
    TimeManager.ExportTimeUseMatrix(Args.JointTimeUseMatrix)
    Return(true)
endMacro


//=========================================================================================================
Macro "Tour Order"(opt)
    vwMem = opt.View
    type = opt.Type
    if Lower(type) = 'joint' then
        idFld = 'HID'
    else
        idFld = 'PerID'

    // Create temporary fields
    obj = CreateObject("Table", vwMem)
    newFlds =  {{FieldName: "NewDepHome", Type: "real", Decimals: 1},
                {FieldName: "NewArrDest", Type: "real", Decimals: 1},
                {FieldName: "RemoveFStop", Type: "Short"},
                {FieldName: "NewDepDest", Type: "real", Decimals: 1},
                {FieldName: "NewArrHome", Type: "real", Decimals: 1},
                {FieldName: "RemoveRStop", Type: "Short"},
                {FieldName: "TourOrder", Type: "Short"},
                {FieldName: "ArrTimePrevTour", Type: "real", Decimals: 1},
                {FieldName: "DepTimeNextTour", Type: "real", Decimals: 1}}
    obj.AddFields({Fields: newFlds})
    
    // Fill tour order after sorting by HHID and departure time
    order = {{idFld, "Ascending"}, {"TourStartTime", "Ascending"}}
    obj.Sort({FieldArray: order})
    vID = obj.(idFld)
    vTourOrder = Vector(vID.Length, "Short",)
    vTourOrder[1] = 1
    c = 1
    for i = 2 to vID.length do
        if vID[i] = vID[i-1] then
            c = c + 1
        else
            c = 1
        
        vTourOrder[i] = c
    end
    obj.TourOrder = vTourOrder
endMacro


//=========================================================================================================
Macro "Select Persons on Tour"(opt)
    abm = opt.abmManager
    vwP = abm.PersonView
    i = opt.TourNumber
    vwT = opt.ToursObj.GetView()
    vwTemp = ExportView(vwT + "|" + opt.TourSet, "MEM", "TempTours", {"HID", "TourOrder", "TourTag"},)

    vwJ = JoinViews("PersonTours", GetFieldFullSpec(vwP, abm.HHIDinPersonView), GetFieldFullSpec(vwTemp, "HID"), )
    exprStr = "if Lower(TourTag) = 'other1' then InJoint_Other1_Tour else if Lower(TourTag) = 'other2' then InJoint_Other2_Tour else InJoint_Shop1_Tour"
    expr = CreateExpression(vwJ, "IsPersonOnTour", exprStr,)
    
    pSet = "__PersonsOnTour" + String(i)
    SetView(vwJ)
    n = SelectByQuery(pSet, "several", "Select * where IsPersonOnTour = 1 and TourOrder = " + String(i),)
    if n = 0 then
        Throw("Unable to select persons on Joint Tour No " + String(i))
    
    CloseView(vwJ)
    CloseView(vwTemp)
    Return(pSet)
endMacro


// ========================================================================================================
Macro "JointStop Scheduling Prep"(opt)
    abm = opt.abmManager
    TimeManager = opt.TimeManager

    TourSpec = {ViewName: opt.ToursObj.GetView(), HHID: "HID", Set: opt.TourSet}
    PersonSpec = {ViewName: abm.PersonView, PersonID: abm.PersonID, Set: opt.PersonSet}   
    
    // From the time manager fill the time after arrival from previous tour and the departure time of the next tour
    // A bit tricky because we are using the tour data as the temp HH data. 
    opts = {PersonSpec: PersonSpec, HHSpec: TourSpec, HHFillField: "ArrTimePrevTour", Metric: "EarliestTime", TimeField: "TourStartTime"}
    TimeManager.FillHHTimeField(opts)

    opts = {PersonSpec: PersonSpec, HHSpec: TourSpec, HHFillField: "DepTimeNextTour", Metric: "LatestTime", TimeField: "TourEndTime"}
    TimeManager.FillHHTimeField(opts)
endMacro


//=========================================================================================================
Macro "JointStop Update TimeManager"(opt)
    abm = opt.abmManager
    TimeManager = opt.TimeManager
    
    vwTempTours = ExportView(opt.ToursObj.GetView() + "|" + opt.TourSet, "MEM", "TempTours", {"HID", "TourStartTime", "TourEndTime"},)
    vwJ = JoinViews("PersonTours", GetFieldFullSpec(abm.PersonView, abm.HHIDinPersonView), GetFieldFullSpec(vwTempTours, "HID"), )
    vwTemp = ExportView(vwJ + "|" + opt.PersonSet, "MEM", "TempPersons", {abm.PersonID, 'HID', 'TourStartTime', 'TourEndTime'},)
    CloseView(vwJ)
    CloseView(vwTempTours)

    TimeManager.UpdateMatrixFromTours({ViewName: vwTemp, PersonID: abm.PersonID, Departure: 'TourStartTime', Arrival: 'TourEndTime'})
    CloseView(vwTemp)
endMacro

// ========================================================================================================
// Calling macros for Solo Tour Stop Models
//==========================================================================================================
Macro "SoloStops Setup"(Args)
    RunMacro("Discretionary Stops Setup", Args, {Type: 'Solo'})
    Return(true)
endMacro


Macro "SoloStops Frequency"(Args)
    RunMacro("Discretionary Stops Freq", Args, {Type: 'Solo'})
    Return(true)
endMacro


Macro "SoloStops Destination"(Args)
    RunMacro("Discretionary Stops Dest", Args, {Type: 'Solo'})
    Return(true)
endMacro


Macro "SoloStops Duration"(Args)
    RunMacro("Discretionary Stops Dur", Args, {Type: 'Solo'})
    Return(true)
endMacro


Macro "SoloStops Scheduling"(Args)
    maxTours = 3 // A person can at most make three solo discretionary tours (O3 or O2S1) on a weekday
    abm = RunMacro("Get ABM Manager", Args)

    // Open Tours file and export to InMemory View
    obj = CreateObject("Table", Args.NonMandatorySoloTours)
    vw = obj.GetView()
    {flds, specs} = GetFields(vw,)
    vwMem = ExportView(vw + "|", "MEM", "SoloToursFile",,)
    obj = null

    objT = CreateObject("Table", vwMem)
    // Assigning tour number order for each person (in chronological order of tour)
    RunMacro("Tour Order", {View: vwMem, Type: 'Solo'})
    
    // Load time manager object
    TimeManager = RunMacro("Get Time Manager", abm)
    TimeManager.LoadTimeUseMatrix({MatrixFile: Args.SoloTimeUseMatrix}) // loading the existing file

    opt = null
    opt.abmManager = abm
    opt.ToursObj = objT
    opt.TimeManager = TimeManager
    pbar = CreateObject("G30 Progress Bar", "Running Intermediate Stops Scheduling for Solo Discretionary Tours (Max 3 tours)", false, maxTours)
    for i = 1 to maxTours do
        opt.TourNumber = i

        // Select records from the tour file
        filter = "TourOrder = " + string(i)
        objT.SelectByQuery({Query: filter, SetName: "__SelectedTours" + String(i)})
        opt.TourSet = "__SelectedTours" + String(i)
        
        // *********** Determine arrival time from previous tour and departure time for next tour
        RunMacro("SoloStop Scheduling Prep", opt)

        // *********** Run Forward Stop Scheduling
        RunMacro("Stop Scheduling Forward", Args, opt)

        // *********** Run Return Stop Scheduling
        RunMacro("Stop Scheduling Return", Args, opt)

        // *********** Update time manager
        RunMacro("SoloStop Update TimeManager", opt)

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()
    obj = null

    // Export the updated in-memory view back to tour table
    ExportView(vwMem + "|", "FFB", Args.NonMandatorySoloTours, flds,)
    //ExportView(vwMem + "|", "FFB", Args.NonMandatorySoloTours, ,)
    CloseView(vwMem)

    // Export Final Time Manager Matrix
    TimeManager.ExportTimeUseMatrix(Args.SoloTimeUseMatrix)
    Return(true)
endMacro


//========================================================================================================
Macro "SoloStop Scheduling Prep"(opt)
    TimeManager = opt.TimeManager

    Spec = {ViewName: opt.ToursObj.GetView(), PersonID: "PerID", Set: opt.TourSet}
    
    // From the time manager fill the time after arrival from previous tour and the departure time of the next tour
    opts = {PersonSpec: Spec, PersonFillField: "ArrTimePrevTour", Metric: "EarliestTime", TimeField: "TourStartTime"}
    TimeManager.FillPersonTimeField(opts)

    opts = {PersonSpec: Spec, PersonFillField: "DepTimeNextTour", Metric: "LatestTime", TimeField: "TourEndTime"}
    TimeManager.FillPersonTimeField(opts)
endMacro


//=========================================================================================================
Macro "SoloStop Update TimeManager"(opt)
    TimeManager = opt.TimeManager
    vwTempTours = ExportView(opt.ToursObj.GetView() + "|" + opt.TourSet, "MEM", "TempTours", {"PerID", "TourStartTime", "TourEndTime"},)
    TimeManager.UpdateMatrixFromTours({ViewName: vwTempTours, PersonID: "PerID", Departure: 'TourStartTime', Arrival: 'TourEndTime'})
    CloseView(vwTempTours)
endMacro


//==========================================================================================================
// File contains common macros for Solo and Joint discretionary intermediate stops
// Stops Setup
Macro "Discretionary Stops Setup"(Args, Opts)
    type = Opts.Type
    file = Args.("NonMandatory" + type + "Tours")
    objT = CreateObject("Table", file)

    flds = {{FieldName: "StopsChoice",          Type: "String", Width: 3, Description: "String of the form A_B, where A is 1 if stop in forward leg, B is 1 if stop in return leg"},
            {FieldName: "ForwardStop",          Type: "Short", Description: "Is there an intermediate stop on forward leg?"},
            {FieldName: "StopForwardTAZ",       Type: "Integer", Description: "Destination TAZ of intermediate stop on forward leg"},
            {FieldName: "ForwardStopDurChoice", Type: "String", Description: "Duration interval of intermediate stop on forward leg"}, 
            {FieldName: "ForwardStopDuration",  Type: "Integer", Description: "Duration of intermediate stop on forward leg in minutes"}, 
            {FieldName: "ForwardStopArrival",   Type: "Integer", Description: "Arrival at the stop location on forward leg"}, 
            {FieldName: "ForwardStopDeparture", Type: "Integer", Description: "Departure from stop location on return leg"}, 
            {FieldName: "ForwardStopDeltaTT",   Type: "Real", Decimals: 1, Description: "Excess or Delta TT due to the forward intermediate stop"},
            {FieldName: "ReturnStop",           Type: "Short", Description: "Is there an intermediate stop on return leg?"}, 
            {FieldName: "StopReturnTAZ",        Type: "Integer", Description: "Destination TAZ of intermediate stop on return leg"}, 
            {FieldName: "ReturnStopDurChoice",  Type: "String", Width: 12, Description: "Duration interval of intermediate stop on return leg"}, 
            {FieldName: "ReturnStopDuration",   Type: "Integer", Description: "Duration of intermediate stop on return leg in minutes"},
            {FieldName: "ReturnStopArrival",    Type: "Integer", Description: "Arrival at the stop location on return leg"}, 
            {FieldName: "ReturnStopDeparture",  Type: "Integer", Description: "Departure from stop location on return leg"},
            {FieldName: "ReturnStopDeltaTT",    Type: "Real", Decimals: 1, Description: "Excess or Delta TT due to the return intermediate stop"}
           }
    objT.AddFields({Fields: flds})
endMacro


//==========================================================================================================
// Stops Frequency Main
Macro "Discretionary Stops Freq"(Args, Opts)
    type = Opts.Type // 'Solo' or 'Joint'
    file = Args.("NonMandatory" + type + "Tours") // Args.NonMandatorySoloTours or Args.NonMandatoryJointTours
    abm = RunMacro("Get ABM Manager", Args)

    purps = {"Other", "Shop"}
    for p in purps do
        RunMacro("Discretionary Stops Freq Eval", Args, {Type: type, Purpose: p, abmManager: abm})
    end

    // Fill in fields after run
    objT = CreateObject("Table", file)
    v = objT.StopsChoice
    arrF = v2a(v).Map(do (f) Return (s2i(Left(f,1))) end)
    arrR = v2a(v).Map(do (f) Return (s2i(Right(f,1))) end)
    
    vecsSet = null
    vecsSet.ForwardStop = a2v(arrF)
    vecsSet.ReturnStop = a2v(arrR)
    objT.SetDataVectors({FieldData: vecsSet})
endMacro


//==========================================================================================================
Macro "Discretionary Stops Freq Eval"(Args, FOpts)
    // Inputs
    type = FOpts.Type
    p = FOpts.Purpose
    abm = FOpts.abmManager
    
    // Derived Variables
    utilFn = Args.(type + "StopFreq" + p + "Utility") // e.g. Args.JointStopFreqOtherUtility
    filter = printf("TourPurpose = '%s'", {p})
    modelName = p + " " + type + " Stops Frequency"
    modelFile = type + "StopFreq_" + p + ".mdl"

    // Join Tours Table to HH File
    toursTable = Args.("NonMandatory" + type + "Tours")
    objT = CreateObject("Table", toursTable)
    if Lower(type) = "joint" then
        vwJ = JoinViews("ToursData", GetFieldFullSpec(objT.GetView(), "HID"), GetFieldFullSpec(abm.HHView, abm.HHID),)
    else
        vwJ = JoinViews("ToursData", GetFieldFullSpec(objT.GetView(), "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID),)
    
    // Run Model
    obj = CreateObject("PMEChoiceModel", {ModelName: modelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + modelFile
    obj.AddTableSource({SourceName: "ToursData", View: vwJ, IDField: "TourID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimOP, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "ToursData", Filter: filter, OField: "HTAZ", DField: "Destination"})
    obj.AddUtility({UtilityFunction: utilFn})
    obj.AddOutputSpec({ChoicesField: "StopsChoice"})
    obj.ReportShares = 1
    obj.RandomSeed = 6299987 + 10*StringLength(type) + StringLength(p)
    ret = obj.Evaluate()
    if !ret then
        Throw("Running '" + type + " Stop Frequency' model for " + p + " failed.")
    Args.(modelName + " Spec") = CopyArray(ret)
    obj = null
    
    if !FOpts.LeaveDataOpen then
        CloseView(vwJ)
endMacro


//==========================================================================================================
Macro "Discretionary Stops Dest"(Args, Opts)
    type = Opts.Type // type = 'Solo' or 'Joint'
    tourFile = Args.("NonMandatory" + type + "Tours")
    utilityFn = Args.(type + "StopDestinationUtility")
    
    izMatrix = Args.IZMatrix
    mobjIZ = CreateObject("Matrix", izMatrix)
    mobjIZ.AddCores({"Ones"})
    mobjIZ.Ones := 1
    mobjIZ = null

    // Time period definitions
    timePeriods = Args.TimePeriods
    amStart = timePeriods.AM.StartTime
    amEnd = timePeriods.AM.EndTime
    pmStart = timePeriods.PM.StartTime
    pmEnd = timePeriods.PM.EndTime

    // Compute size variable field. Add to TAZDemographics output table
    objD = CreateObject("Table", Args.DemographicOutputs)
    obj4D = CreateObject("Table", Args.AccessibilitiesOutputs)
    sizeVarFld = type + "StopSizeVar"
    newFlds = {{FieldName: sizeVarFld, Type: "real", Width: 12, Decimals: 2}}
    objD.AddFields({Fields: newFlds})
    objJ = objD.Join({Table: obj4D, LeftFields: {"TAZ"}, RightFields: {"TAZID"}})
    
    opt = null
    opt.TableObject = objJ
    opt.Equation = Args.(type + "StopSizeVar")
    opt.FillField = sizeVarFld
    opt.ExponentiateCoeffs = 1
    RunMacro("Compute Size Variable", opt)
    objJ = null
    objD = null
    obj4D = null

    // Run Destination Choice
    objT = CreateObject("Table", tourFile)
    vwT = objT.GetView()
    spec = {ToursView: vwT}
    
    dirs = {"Forward", "Return"}
    periods = {"AM", "PM", "OP"}
    pbar = CreateObject("G30 Progress Bar", "Running " + type + " Intermediate Stops Destination Choice for combination (Forward, Return) and (AM, PM, OP)", false, 6)
    for dir in dirs do // Tour direction loop
        // Dir specific data
        if dir = "Forward" then do
            ODInfo.Origin = "HTAZ"
            ODInfo.Destination = "Destination"
            depFld = "TourStartTime"
        end
        else do
            ODInfo.Origin = "Destination"
            ODInfo.Destination = "HTAZ"
            depFld = "DestDepTime"
        end
        spec.StopFilter = printf("%sStop > 0", {dir}) // e.g. ForwardStop > 0
        spec.ODInfo = ODInfo

        amQry = printf("(%s >= %s and %s < %s)", {depFld, String(amStart), depFld, String(amEnd)})
        pmQry = printf("(%s >= %s and %s < %s)", {depFld, String(pmStart), depFld, String(pmEnd)})
        exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
        depPeriod = CreateExpression(vwT, "DepPeriod", exprStr,)

        // Period Loop
        for period in periods do
            // Set Time Period Filter
            spec.PeriodFilter = printf("DepPeriod = '%s'", {period})
            skimFile = Args.("HighwaySkim" + period)

            // Run Intermediate Travel Times and Destination Choice
            // Calculate delta TT matrix
            deltaTT = GetTempPath() + "DeltaTT_" + dir + "_" + period + ".mtx"
            spec.DeltaSkim = deltaTT
            spec.MatrixSpec = {File: skimFile, Core: "Time", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
            spec.MatrixSpecR = null
            spec.OutputCoreName = "DeltaTT"
            ret = RunMacro("Calculate Delta TT", Args, spec)
            if ret = 2 then continue // No records for delta TT calculation. Move on to next period.

            // Calculate delta dist matrix
            deltaDist = GetTempPath() + "DeltaDist_" + dir + "_" + period + ".mtx"
            spec.DeltaSkim = deltaDist
            spec.MatrixSpec = {File: skimFile, Core: "Distance", RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"}
            spec.MatrixSpecR = null
            spec.OutputCoreName = "DeltaDist"
            ret = RunMacro("Calculate Delta TT", Args, spec)
            if ret = 2 then continue
            
            // Calculate delta StopIsDest mtx
            deltaStop = GetTempPath() + "DeltaStopSameAsDest_" + dir + "_" + period + ".mtx"
            spec.DeltaSkim = deltaStop
            spec.MatrixSpec = {File: izMatrix, Core: "Ones", RowIndex: "TAZ", ColIndex: "TAZ"}
            spec.MatrixSpecR = {File: izMatrix, Core: "IZ", RowIndex: "TAZ", ColIndex: "TAZ"}
            spec.OutputCoreName = "StopSameAsDest"
            ret = RunMacro("Calculate Delta TT", Args, spec)
            if ret = 2 then continue

            // Run Dest Choice
            tag = type + "Stops_" + dir + "_" + period
            filter = printf("(%s) and (%s)", {spec.StopFilter, spec.PeriodFilter})
            obj = CreateObject("PMEChoiceModel", {ModelName: tag})
            obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + ".dcm"
            obj.AddTableSource({SourceName: "ToursData", View: vwT, IDField: "TourID"})
            obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
            obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
            obj.AddMatrixSource({SourceName: "AutoSkim", File: skimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
            obj.AddMatrixSource({SourceName: "Intrazonal", File: izMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
            obj.AddMatrixSource({SourceName: "DeltaTT", File: deltaTT, PersonBased: 1})
            obj.AddMatrixSource({SourceName: "DeltaDist", File: deltaDist, PersonBased: 1})
            obj.AddMatrixSource({SourceName: "StopSameAsDestMtx", File: deltaStop, PersonBased: 1})
            obj.AddPrimarySpec({Name: "ToursData", Filter: filter, OField: ODInfo.Origin})
            obj.AddUtility({UtilityFunction: utilityFn})
            obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
            obj.AddSizeVariable({Name: "TAZData", Field: sizeVarFld})
            obj.AddOutputSpec({ChoicesField: "Stop" + dir + "TAZ"})
            obj.RandomSeed = 6399971 + 100*StringLength(type) + 10*StringLength(dir) + periods.Position(period)
            ret = obj.Evaluate()
            if !ret then
                Throw("Running 'Stop Location' model failed for: " + type + "_" + dir + "_" + period)

            if pbar.Step() then
                Return()
        end // period loop

        // Now that destinations have been computed, fill in the realized detour travel times
        opt = {ToursObj: objT, Filter: printf("%sStop > 0", {dir}), ODInfo: ODInfo, 
                StopTAZField: "Stop" + dir + "TAZ", ModeField: "Mode", DepTimeField: depFld, 
                OutputField: dir + "StopDeltaTT"}
        RunMacro("Calculate Detour TT", Args, opt)

        // Remove infeasible stops (where stops delta TT is null or more than 45 min)
        filter = printf("%sStop > 0 and (%sStopDeltaTT = null or %sStopDeltaTT > 45)", {dir, dir, dir})
        opt = {TableObject: objT, Filter: filter, Direction: dir}
        RunMacro("Remove NM Infeasible Stops", opt)

        DestroyExpression(GetFieldFullSpec(vwT, depPeriod))
    end // dir loop
    pbar.Destroy()
endMacro


//==========================================================================================================
Macro "Discretionary Stops Dur"(Args, Opts)
    type = Opts.Type // 'Solo' or 'Joint'
    file = Args.("NonMandatory" + type + "Tours")
    objT = CreateObject("Table", file)
    abm = RunMacro("Get ABM Manager", Args)

    purps = {"Other", "Shop"}
    dirs = {"Forward", "Return"}
    
    pbar = CreateObject("G30 Progress Bar", "Running " + type + " Intermediate Stops Duration Choice for combination of (Other, Shop) and (Forward, Return) stop", true, 4)
    for purp in purps do
        for dir in dirs do
            // dir specific data
            freqFld = dir + "Stop"
            choiceIntFld = dir + "StopDurChoice"
            choiceFld = dir + "StopDuration"
            filter = freqFld + " > 0 and TourPurpose = '" + purp + "'"

            // Run Duration choice model
            opt = {abmManager: abm, ToursObj: objT, Type: type, Purpose: purp, Direction: dir, IntegerChoiceField: choiceIntFld, Filter: filter}
            RunMacro("Stops Duration Eval", Args, opt)

            // Simulate Time based on choice interval
            vw = objT.GetView()
            SetView(vw)
            n = SelectByQuery("__Selection", "several", "Select * where " + filter,)
            opt = {ViewSet: vw + "|__Selection", InputField: choiceIntFld, OutputField: choiceFld, AlternativeIntervalInMin: 1}
            RunMacro("Simulate Time", opt)

            if pbar.Step() then
                Return()
        end
    end
    pbar.Destroy()
    objT = null
endMacro


// =========================================================================================================
Macro "Stops Duration Eval"(Args, opt)
    type = opt.Type
    purp = opt.Purpose
    dir = opt.Direction
    choiceIntFld = opt.IntegerChoiceField
    filter = opt.Filter
    abm = opt.abmManager
    
    primaryViewName = type + "ToursData"
    utilFn = Args.(type + "StopDur" + purp + "Utility")
    utilOpts = null
    utilOpts.UtilityFunction = utilFn
    utilOpts.SubstituteStrings = { {"<dir>", dir} }

    // Join Tours Table to HH File
    objT = opt.ToursObj
    if Lower(type) = "joint" then
        vwJ = JoinViews("ToursData", GetFieldFullSpec(objT.GetView(), "HID"), GetFieldFullSpec(abm.HHView, abm.HHID),)
    else
        vwJ = JoinViews("ToursData", GetFieldFullSpec(objT.GetView(), "PerID"), GetFieldFullSpec(abm.PersonHHView, abm.PersonID),)
    
    // Run Duration choice model
    tag = type + "Stops_" + purp + "_" + Left(dir,1)
    modelName = tag + " Duration"
    obj = CreateObject("PMEChoiceModel", {ModelName: modelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + "_Duration.mdl"
    obj.AddTableSource({SourceName: "ToursData", View: vwJ, IDField: "TourID"})
    obj.AddPrimarySpec({Name: "ToursData", Filter: filter})
    obj.AddUtility(utilOpts)
    obj.AddOutputSpec({ChoicesField: choiceIntFld})
    obj.ReportShares = 1
    obj.RandomSeed = 6499991 + 100*StringLength(type) + 10*StringLength(purp) + StringLength(dir)
    ret = obj.Evaluate()
    if !ret then
        Throw("Running '" + type + " Stop Duration' model failed for: " + purp + "_" + dir)
    Args.(modelName + " Spec") = CopyArray(ret)
    obj = null

    if !opt.LeaveDataOpen then
        CloseView(vwJ)
endMacro


//=========================================================================================================
// Forward Joint/Solo Stop Scheduling Logic
Macro "Stop Scheduling Forward"(Args, opt)
    toursObj = opt.ToursObj
    vwT = toursObj.GetView()
    tourNo = opt.TourNumber
    MasterFilter = "ForwardStop = 1 and TourOrder = " + String(tourNo)
    nF = toursObj.SelectByQuery({SetName: "__ForwardStops", Query: MasterFilter})
    if nF > 0 then do
        flds = {"ForwardStopDeltaTT", "ForwardStopDuration", "TourStartTime", 
                "DestArrTime", "ArrTimePrevTour", "ActivityDuration", "RemoveFStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ForwardStopDeltaTT + vecs.ForwardStopDuration
        vDesiredStart = vecs.TourStartTime - vStopTime  // Desired start time if there is no overlap onto previous tour
        
        // Check if early start overlaps with previous tour
        vOverlapPrevTour = if (vDesiredStart < vecs.ArrTimePrevTour + 15) then 1 else 0
        
        // And set delay to diff of new constrained start time and the desired start time
        vDelay = if vOverlapPrevTour = 1 then (vecs.ArrTimePrevTour + 15) - vDesiredStart else 0
        vNewActDur = vecs.ActivityDuration - vDelay // Realized activity duration due to the delay
        vRemoveStop = if (vNewActDur/vecs.ActivityDuration < 0.33 or vNewActDur < 5) then 1 else 0

        vecsSet = null
        vecsSet.TourStartTime = if vRemoveStop <> 1 then vDesiredStart + vDelay else vecs.TourStartTime
        vecsSet.DestArrTime = if vRemoveStop <> 1 then vecs.DestArrTime + vDelay else vecs.DestArrTime
        vecsSet.RemoveFStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})

        // Calculate arrival time at stop, departure time at stop and arrival time at main destination
        fillSpec = {View: vwT, OField: "HTAZ", DField: "StopForwardTAZ", FillField: "TimeToStopF", 
                    Filter: MasterFilter, ModeField: "Mode", DepTimeField: "TourStartTime"}
        RunMacro("Fill Travel Times", Args, fillSpec)

        toursObj.ChangeSet({SetName: "__ForwardStops"})

        // Get Stop Arrival and dep Time
        toursObj.ForwardStopArrival = toursObj.TourStartTime + toursObj.TimeToStopF
        toursObj.ForwardStopDeparture = toursObj.ForwardStopArrival + toursObj.ForwardStopDuration

        // Get travel time from stop to main destination
        fillSpec = {View: vwT, OField: "StopForwardTAZ", DField: "Destination", FillField: "TimeFromStopF", 
                    Filter: MasterFilter, ModeField: "Mode", DepTimeField: "ForwardStopDeparture"}
        RunMacro("Fill Travel Times", Args, fillSpec)

        toursObj.ChangeSet({SetName: "__ForwardStops"})

        // Get arrival at main destination
        toursObj.DestArrTime = toursObj.ForwardStopDeparture + toursObj.TimeFromStopF
    end

    // Remove infeasible forward stop
    filter = "RemoveFStop = 1 and TourOrder = " + String(tourNo)
    opt = {TableObject: toursObj, Filter: filter, Direction: "Forward"}
    RunMacro("Remove NM Infeasible Stops", opt)
endMacro


//=========================================================================================================
// Return Joint/Solo Stop Scheduling Logic
Macro "Stop Scheduling Return"(Args, opt)
    toursObj = opt.ToursObj
    vwT = toursObj.GetView()
    tourNo = opt.TourNumber
    MasterFilter = "ReturnStop = 1 and TourOrder = " + String(tourNo)
    nR = toursObj.SelectByQuery({SetName: "__ReturnStops", Query: MasterFilter})
    if nR > 0 then do
        flds = {"ReturnStopDeltaTT", "ReturnStopDuration", "TourEndTime", "DestDepTime", 
                "DepTimeNextTour", "ActivityDuration", "RemoveRStop"}
        vecs = toursObj.GetDataVectors({FieldNames: flds})
        vStopTime = vecs.ReturnStopDeltaTT + vecs.ReturnStopDuration
        vLateArr = vecs.TourEndTime + vStopTime
        vOverlapNextTour = if (vLateArr > vecs.DepTimeNextTour) then 1 else 0
        vDelay = if vOverlapNextTour = 1 then vLateArr - vecs.DepTimeNextTour else 0
        vNewActDur = vecs.ActivityDuration - vDelay
        vRemoveStop = if (vNewActDur/vecs.ActivityDuration < 0.33 or vNewActDur < 5) then 1 else 0
        
        vecsSet = null
        vecsSet.DestDepTime = if vRemoveStop <> 1 then vecs.DestDepTime - vDelay else vecs.DestDepTime
        vecsSet.TourEndTime = if vRemoveStop <> 1 then vLateArr - vDelay else vecs.TourEndTime
        vecsSet.RemoveRStop = vRemoveStop
        toursObj.SetDataVectors({FieldData: vecsSet})

        // Calculate arrival time at stop, departure time at stop and arrival time at main destination
        fillSpec = {View: vwT, OField: "Destination", DField: "StopReturnTAZ", FillField: "TimeToStopR", 
                    Filter: MasterFilter, ModeField: "Mode", DepTimeField: "DestDepTime"}
        RunMacro("Fill Travel Times", Args, fillSpec)

        toursObj.ChangeSet({SetName: "__ReturnStops"})

        // Get Stop Arrival and dep Time
        toursObj.ReturnStopArrival = toursObj.DestDepTime + toursObj.TimeToStopR
        toursObj.ReturnStopDeparture = toursObj.ReturnStopArrival + toursObj.ReturnStopDuration

        // Get travel time from stop to main destination
        fillSpec = {View: vwT, OField: "StopReturnTAZ", DField: "HTAZ", FillField: "TimeFromStopR", 
                    Filter: MasterFilter, ModeField: "Mode", DepTimeField: "ReturnStopDeparture"}
        RunMacro("Fill Travel Times", Args, fillSpec)

        toursObj.ChangeSet({SetName: "__ReturnStops"})

        // Get arrival at main destination
        toursObj.TourEndTime = toursObj.ReturnStopDeparture + toursObj.TimeFromStopR
    end
    
    filter = "RemoveRStop = 1 and TourOrder = " + String(tourNo)
    opt = {TableObject: toursObj, Filter: filter, Direction: "Return"}
    RunMacro("Remove NM Infeasible Stops", opt)
endMacro


// Macro that removes the intermediate stop (choice, duration, destination info) if not feasible.
// This happens because the stop may not be feasible given the mode (such as "Walk")
Macro "Remove NM Infeasible Stops"(opt)
    obj = opt.TableObject
    dir = opt.Direction
    n = obj.SelectByQuery({Query: opt.Filter, SetName: "__Remove"})

    if n > 0 then do
        AppendToLogFile(1, opt.Message + String(n) + " Intermediate " + dir + " stops removed due to schedule constraints.")
        v = Vector(n, "Long",)
        vS = Vector(n, "String",)
        vecsSet = null
        vecsSet.(dir + "Stop") = nz(v)
        vecsSet.("Stop" + dir + "TAZ") = v
        vecsSet.(dir + "StopDurChoice") = vS
        vecsSet.(dir + "StopDuration") = v
        vecsSet.(dir + "StopDeltaTT") = v
        vecsSet.(dir + "StopArrival") = v
        vecsSet.(dir + "StopDeparture") = v
        vStopsChoice = obj.StopsChoice
        if dir = 'Forward' then
            vecsSet.StopsChoice = "0_" + Right(vStopsChoice, 1)
        else
            vecsSet.StopsChoice = Left(vStopsChoice, 1) + "_0"
        obj.SetDataVectors({FieldData: vecsSet})
    end
    obj.ChangeSet()
endMacro
