// ******************************* Mandatory Mode Models *************************************************************
// *** Determine mode for the mandatory tours
// All tours of the same type by the same person are assumed to have the same mode. Example, both work tours for a person have the same mode
Macro "Mandatory Mode Choice"(Args)
    // Preprocess
    abm = RunMacro("Get ABM Manager", Args)
    
    // Run preprocess and initial vehicle priorities
    RunMacro("Mode Choice Preprocess", Args, abm)
    
    // For the Work MC calibration, it is enough to first run the mode choice upto this point
    if Args.WorkMC_Calibration then
        Return(true)

    pbar = CreateObject("G30 Progress Bar", "Mandatory Mode Choice", true, 5)
    // Module for full time workers making work tours
    spec = {abmManager: abm, 
            Type: 'Work',
            Category: 'FullTimeWorker',
            Filter: 'WorkerCategory = 1 and TravelToWork = 1',
            Alternatives: Args.WorkModes, 
            Utility: Args.WorkModeUtility,
            Availability: Args.WorkModeAvailability,
            LocationField: "WorkTAZ",
            ChoiceField: 'WorkMode',
            ChoiceCodeField: 'WorkModeCode',
            RandomSeed: 3099997}
    RunMacro("Mode Choice Module", Args, spec)
    if pbar.Step() then
        Return()

    // Module for part time workers and univ students making work tours
    spec = {abmManager: abm, 
            Type: 'Work',
            Category: 'PartTimeWorker',
            Filter: 'WorkerCategory <> 1 and TravelToWork = 1',
            Alternatives: Args.WorkModes,  
            Utility: Args.WorkModeUtility,
            Availability: Args.WorkModeAvailability,
            LocationField: "WorkTAZ",
            ChoiceField: 'WorkMode',
            ChoiceCodeField: 'WorkModeCode',
            RandomSeed: 3199997}
    RunMacro("Mode Choice Module", Args, spec)
    if pbar.Step() then
        Return()

    // For the Univ MC calibration, it is enough to first run the mode choice upto this point.
    // Note that the Work MC decisions are complete (which influences how many vehicles are left for Univ students)
    if Args.UnivMC_Calibration then
        Return(true)

    // Module for anyone making univ tours
    spec = {abmManager: abm, 
            Type: 'Univ',
            Category: 'University student',
            Filter: 'AttendUniv = 1',
            Alternatives: Args.UnivModes,  
            Utility: Args.UnivModeUtility,
            Availability: Args.UnivModeAvailability,
            LocationField: "UnivTAZ",
            ChoiceField: 'UnivMode',
            ChoiceCodeField: 'UnivModeCode',
            RandomSeed: 3299969}
    ret = RunMacro("Mode Choice Module", Args, spec)
    if pbar.Step() then
        Return()

    // For the School MC (Forward trip) calibration, it is enough to first run the mode choice upto this point.
    // Note that the Work and HigherEd MC decisions are complete (which influences the carpool eligibility)
    if Args.SchoolMCF_Calibration then
        Return(true)

    // Module for school tours (Forward Mode)
    spec = {abmManager: abm, 
            Type: 'School',
            Category: 'School',
            Filter: 'AttendSchool = 1',
            Alternatives: Args.SchoolModes, 
            Utility: Args.SchoolModeFUtility,
            Availability: Args.SchModeFAvailability,
            LocationField: "SchoolTAZ",
            ChoiceField: 'SchoolForwardMode',
            ChoiceCodeField: 'SchoolForwardModeCode',
            RandomSeed: 3399997}
   RunMacro("Mode Choice Module", Args, spec)

    // For the School MC (Return trip) calibration, it is enough to first run the mode choice upto this point.
    // Note that the Work and HigherEd MC decisions are complete (which influences the carpool eligibility)
    // The School forward mode is also run. This is required to ensured that students who biked to school return by bike.
    if Args.SchoolMCR_Calibration then
        Return(true)

    // Module for school tours (Return Mode)
    spec = {abmManager: abm, 
            Type: 'School',
            Category: 'School',
            Filter: "(AttendSchool = 1 and SchoolForwardMode <> 'Bike' and SchoolForwardMode <> 'DriveAlone')",
            Direction: 'Return',
            Alternatives: Args.SchoolModes, 
            Utility: Args.SchoolModeRUtility,
            Availability: Args.SchModeRAvailability,
            LocationField: "SchoolTAZ",
            ChoiceField: 'SchoolReturnMode',
            ChoiceCodeField: 'SchoolReturnModeCode',
            RandomSeed: 3499999}
    RunMacro("Mode Choice Module", Args, spec)
    if pbar.Step() then
        Return()

    RunMacro("Daycare Mode Choice", abm)
    
    RunMacro("School MC Postprocess", abm)

    RunMacro("Export ABM Data", Args)
    pbar.Destroy()
    return(true)
endMacro


Macro "Mode Choice Preprocess"(Args, abm)
    vwP = abm.PersonView
    
    // Determine priority order to determine who is allocated to a vehicle in the HH
    // Note some ground rules
    // 1. No person can attend both school and university
    // 2. No school age student (Age <= 18) can also be a full time worker. They may be part time workers but have no work from home option
    // 3. Workers may not travel on given day because they are either not attending work or they are working from home
    expr = "if Age < 15 or License <> 1 then 100 " + // Kids + Unlicensed folks
           "else if WorkerCategory = 1 and AttendUniv = 1 and TravelToWork = 1 then 1 " + // Full time workers traveling to work on given day and also attending university
           "else if WorkerCategory = 2 and AttendUniv = 1 and TravelToWork = 1 then 2 " + // Part time workers traveling to work on given day and also attending university
           "else if WorkerCategory = 1 and AttendUniv <> 1 and TravelToWork = 1 then 3 " +  // Full time workers traveling to work on given day but not attending university
           "else if WorkerCategory = 2 and AttendUniv <> 1 and TravelToWork = 1 then 4 " +  // Part time workers traveling to work on given day but not attending university (they may attend school however)
           "else if WorkerCategory <= 2 and AttendUniv = 1 and TravelToWork <> 1 then 5 " +  // Working university students not traveling to work on given day
           "else if WorkerCategory = null and AttendUniv = 1 then 6 " +  // Non-working university students     
           "else if WorkerCategory = 1 and TravelToWork <> 1 then 7 " +    // Full time workers not attending univ and not traveling on given day
           "else if WorkerCategory = 2 and AttendSchool <> 1 and TravelToWork <> 1 then 8 " +    // Part time workers not attending school, not attending univ and not traveling on given day
           "else if WorkerCategory = null and (Age > 18 and Age < 65) then 9 " +    // Non-Seniors who are neither working nor attending university
           "else if AttendSchool = 1 then 10 " +   // Students with license and attending school
           "else if WorkerCategory = null and Age >= 65 then 11 " +  // Seniors who are neither working nor attending university
           "else if AttendSchool <> 1 then 12"     // Anyone over 15 years old with license but not attending school, univ or work
    expr1 = CreateExpression(vwP, "Temp", expr,)
    abm.ActivatePersonSet(null)
    v = abm.[Person.Temp]
    vZero = Vector(v.Length, "Short", {{"Constant", 0}})
    abm.SetPersonVectors({VehiclePriority: v, VehicleUsed: vZero, VehicleAvail: vZero})
    DestroyExpression(GetFieldFullSpec(vwP, expr1))

    // Fill vehicles remaining field in HH layer with total number of vehicles
    abm.ActivateHHSet()
    abm.[HH.VehiclesRem] = abm.[HH.Vehicles]
endMacro


Macro "Mode Choice Module"(Args, spec)
        pbarMC = CreateObject("G30 Progress Bar", "Mode Choice Preprocess for " + spec.Category, true, 3)
        // 1. MC Preparation Macro
        // Vehicle Allocation: Not required for school mode since there is no DA
        if Lower(spec.Type) <> 'school' then
            RunMacro("Vehicle Allocation", spec)
        
        if pbarMC.Step() then
            Return()

        // 1a. Run Carpool eligibility for school mode choice (need to do this only once)
        pbarMC.SetMessage("Carpool Eligility for school mode")
        if spec.Type = 'School' and spec.Direction <> 'Return' then
            RunMacro("School Carpool Eligibility", spec.abmManager)
        
        // 2. Run MC
        pbarMC.SetMessage("Mode Choice Evaluation for " + spec.Category)
        RunMacro("Evaluate Mode Choice", Args, spec)
        if pbarMC.Step() then
            Return()

        // 3. MC Post Process
        // Macro that looks at choices made and:
        // 1. Reallocates vehicles to other members in the HH based on unused vehicles
        // 2. Compute mode and departure time specific travel times from home to mandatory destination and back.
        pbarMC.SetMessage("Mode Choice PostProcess for " + spec.Category)
        RunMacro("Mode Choice PostProcess", Args, spec)
        pbarMC.Destroy()
endMacro


// Allocate cars among members in the HH
Macro "Vehicle Allocation"(spec)
    abm = spec.abmManager

    baseFilter = "VehiclePriority < 100 and VehicleUsed <> 1" // The latter condition to ensure that persons who already used a vehicle retain it
    if spec.Filter = null then
        filter = baseFilter
    else
        filter = printf("%s and %s", {spec.Filter, baseFilter})
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})

    ttFld = printf("HomeTo%sTime", {spec.Type})
    
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "Allocate Vehicles", 
                InputFields: {"VehiclesRem"},
                OutputFields: {"VehicleAvail"},
                SortOrder: {{"VehiclePriority", "Ascending"}, {ttFld, "Descending"}, {"Age", "Descending"}}
                }
    abm.Iterate(iterOpts)
endMacro


Macro "Allocate Vehicles"(spec)
    inputVecs = spec.InputVecs
    outputVecs = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    
    // Get number of vehicles remaining
    vRem = inputVecs.VehiclesRem
    nRem = vRem[startIdx]

    // Since data is sorted for this HH by priority order, allocate remaining vehicles to people
    for i = startIdx to endIdx do
        if i < startIdx + nRem then
            outputVecs.VehicleAvail[i] = 1
        else
            outputVecs.VehicleAvail[i] = 0
    end
endMacro


// Mode Choice NLM macro
// For the given tour type, loop over the three TOD: AM, PM and OP and use records where departure period matches the TOD
Macro "Evaluate Mode Choice"(Args, spec)
    abm = spec.abmManager
    type = spec.Type
    category = spec.Category
    filter = spec.Filter                    
    destField = spec.LocationField
    direction = spec.Direction // only for school
    purp = type 
    if (Lower(type) = 'work' or Lower(type) = 'univ') then // Use first work or univ tour
        purp = type + "1"

    // Obtain util spec and include availabilities
    utilSpec = {UtilityFunction: spec.Utility, AvailabilityExpressions: spec.Availability}

    // Create TOD expression for each of the periods
    timePeriods = Args.TimePeriods
    amStart = timePeriods.AM.StartTime
    amEnd = timePeriods.AM.EndTime
    pmStart = timePeriods.PM.StartTime
    pmEnd = timePeriods.PM.EndTime
    
    vwPHH = abm.PersonHHView
    if direction = 'Return' then
        depTimeExpr = printf("%s_StartTime + %s_Duration", {purp, purp})
    else
        depTimeExpr = printf("%s_StartTime - HomeTo%sTime", {purp, type})
    depTime = CreateExpression(vwPHH, "DepTime", depTimeExpr,)
    timeAllowance = '0'
    amQry = printf("(%s - %s >= %s and %s - %s < %s)", {depTime, timeAllowance, String(amStart), depTime, timeAllowance, String(amEnd)})
    pmQry = printf("(%s - %s >= %s and %s - %s < %s)", {depTime, timeAllowance, String(pmStart), depTime, timeAllowance, String(pmEnd)})
    exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
    depPeriod = CreateExpression(vwPHH, "DepPeriod", exprStr,)

    tods = {"AM", "PM", "OP"}
    for tod in tods do
        todFilter = printf("%s = '%s'", {depPeriod, tod})
        finalFilter = printf("(%s) and (%s)", {filter, todFilter})
        
        // Skip if no records
        set = abm.CreatePersonSet({Filter: finalFilter, UsePersonHHView: 1})
        if set = null or set.Size = 0 then
            continue

        autoSkimFile = Args.HighwaySkim
        ptWalkSkimFile = Args.TransitWalkSkim
        ptDriveSkimFile = Args.TransitDriveSkim
        
        // Run Model and populate results
        tag = category + "_" + tod + "_Mode" + direction
        obj = CreateObject("PMEChoiceModel", {ModelName: tag + " Tour Mode"})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + tag + ".mdl"
        obj.AddAlternatives({AlternativesTree: spec.Alternatives})
        obj.AddTableSource({SourceName: "PersonHH", View: vwPHH, IDField: abm.PersonID})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: autoSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "PTWalkSkim", File: ptWalkSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "PTDriveSkim", File: ptDriveSkimFile, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "WalkSkim", File: Args.WalkSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "BikeSkim", File: Args.BikeSkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        if direction = 'Return' then
            obj.AddPrimarySpec({Name: "PersonHH", Filter: finalFilter, OField: destField, DField: "TAZID"})
        else
            obj.AddPrimarySpec({Name: "PersonHH", Filter: finalFilter, OField: "TAZID", DField: destField})     
        obj.AddUtility(utilSpec)
        obj.AddOutputSpec({ChoicesField: spec.ChoiceField})
        obj.ReportShares = 1
        obj.RandomSeed = spec.RandomSeed
        ret = obj.Evaluate()
        if !ret then
            Throw("Running '" + tag + " TOD' choice model failed.")
        Args.(tag + " Spec") = CopyArray(ret) // For calibration purposes
    end
    DestroyExpression(GetFieldFullSpec(vwPHH, depPeriod))
    DestroyExpression(GetFieldFullSpec(vwPHH, depTime))
    Return(ret)
endMacro


// Macro that forms final carpool groups among HH members, designates carpool drivers and evaluates vehicle inventory
// Macro also updates estimate of travel time from home to main destination after considering the mode.
Macro "Mode Choice PostProcess"(Args, spec)
    // Recompute the number of vehicles remaining (for the next segment) as a result of the model choices
    if Lower(spec.Type) <> 'school' then
        RunMacro("Vehicle Inventory", spec)

    // Update estimate of time from home to main destination (and from main destination to home) now that mode is determined
    RunMacro("Update Home to Dest TT", Args, spec)
    RunMacro("Update Dest to Home TT", Args, spec)

    // Attach Mode Codes
    codeMap = {DriveAlone: 1, Carpool: 2, Walk: 3, Bike: 4, PTWalk: 5, PTDrive: 6, Other: 7, SchoolBus: 8}
    abm = spec.abmManager
    set = abm.CreatePersonSet({Filter: spec.Filter, Activate: 1})
    inputFld = spec.ChoiceField
    outputFld = spec.ChoiceCodeField
    vecs = abm.GetPersonVectors({inputFld})
    arrMode = v2a(vecs.(inputFld))
    arrModeCode = arrMode.Map(do (f) Return(codeMap.(f)) end)
    vecsSet.(outputFld) = a2v(arrModeCode)
    abm.SetPersonVectors(vecsSet)

    // Run additional post processing for school
    //if spec.Type = 'School' and spec.Direction = 'Return' then
    //    RunMacro("School MC Postprocess", spec.abmManager)
endMacro


/*
    Simple daycare mode choice
    If Dropoff/Pickup person available, mode is 'Carpool' else mode is 'SchoolBus' (Private transportation)
*/
Macro "Daycare Mode Choice"(abm)
    set = abm.CreatePersonSet({Filter: 'AttendDaycare = 1', Activate: 1})
    vecs = abm.GetPersonVectors({"DropoffPersonID", "PickupPersonID"})
    
    vecsSet = null
    vecsSet.SchoolForwardMode = if vecs.DropoffPersonID = null then 'SchoolBus' else 'Carpool'
    vecsSet.SchoolReturnMode = if vecs.PickupPersonID = null then 'SchoolBus' else 'Carpool'
    vecsSet.SchoolForwardModeCode = if vecs.DropoffPersonID = null then 8 else 2
    vecsSet.SchoolReturnModeCode = if vecs.PickupPersonID = null then 8 else 2
    abm.SetPersonVectors(vecsSet)
endmacro


Macro "Vehicle Inventory"(opt)
    abm = opt.abmManager
    vwP = abm.PersonView
    filter = opt.Filter
    choiceField = opt.ChoiceField

    if abm = null or choiceField = null or filter = null then
        Throw("Missing inputs to \'Vehicle Inventory\' macro")

    // Expression: VUsed = 1 If vehicle is required based on chosen mode
    expr = "if " + choiceField + " = 'DriveAlone' then 1 else 0"
    expr1 = CreateExpression(vwP, "VUsed", expr,)
    
    // Expression for additional vehicle used.
    // AddnlVehUsed = 1 If vehicle is required and person already has not used up a vehicle
    expr = "if VehicleUsed <> 1 and VUsed = 1 then 1 else 0"
    expr2 = CreateExpression(vwP, "AddnlVehUsed", expr,)

    // Aggregate additional vehciles used for this step at the HH level and update remaining vehicles field
    aggSpec = printf("(%s).AddnlVehUsed.Sum", {filter})
    aggOpts.Spec = {{VehiclesUsed: aggSpec, DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)
    abm.[HH.VehiclesRem] = abm.[HH.VehiclesRem] - abm.[HH.VehiclesUsed]
    
    // Update vehicle used field in Person file
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})
    abm.[Person.VehicleUsed] = if abm.[Person.VUsed] = 1 or abm.[Person.VehicleUsed] = 1 then 1 else 0

    DestroyExpression(GetFieldFullSpec(vwP, "VUsed"))
    DestroyExpression(GetFieldFullSpec(vwP, "AddnlVehUsed"))
endMacro


// Macro that updates home to work/univ/school time for persons who chose NonAuto mode
Macro "Update Home to Dest TT"(Args, opt)
    type = opt.Type
    if (opt.Direction = 'Return') and (Lower(type) = 'school')
        then Return()
    abm = opt.abmManager
    vwPHH = abm.PersonHHView

    purp = opt.Type 
    if (Lower(type) = 'work' or Lower(type) = 'univ') then // Use first work or univ tour
        purp = type + "1"

    // Get approx departure time
    fillField = printf("HomeTo%sTime", {type})
    depTimeExpr = printf("%s_StartTime - HomeTo%sTime", {purp, type})
    depTimeFld = CreateExpression(vwPHH, "DepTime", depTimeExpr,)
    fillSpec = {View: vwPHH, OField: "TAZID", DField: opt.LocationField, FillField: fillField, 
                Filter: opt.Filter, ModeField: opt.ChoiceField, DepTimeField: depTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwPHH, depTimeFld))
endMacro


// Macro that updates home to work/univ/school time for persons who chose NonAuto mode
Macro "Update Dest to Home TT"(Args, opt)
    type = opt.Type
    if (opt.Direction <> 'Return') and (Lower(type) = 'school')
        then Return()
    abm = opt.abmManager
    vwPHH = abm.PersonHHView

    purp = opt.Type 
    if (Lower(type) = 'work' or Lower(type) = 'univ') then // Use first work or univ tour
        purp = type + "1"

    // Get approx departure time
    fillField = printf("%sToHomeTime", {type})
    depTimeExpr = printf("%s_StartTime + %s_Duration", {purp, purp})
    depTimeFld = CreateExpression(vwPHH, "DepTime", depTimeExpr,)
    fillSpec = {View: vwPHH, OField: opt.LocationField, DField: "TAZID", FillField: fillField, 
                Filter: opt.Filter, ModeField: opt.ChoiceField, DepTimeField: depTimeFld}
    RunMacro("Fill Travel Times", Args, fillSpec)
    DestroyExpression(GetFieldFullSpec(vwPHH, depTimeFld))
endMacro


/*  Postprocess for school
    Called after the school return mode choice
    1. Set return mode for those who chose to bike to school on the forward trip
    2. Aggregate person fields to determine the # kids dropped off and picked up by each HH
    3. Update person fields determine the actual # kids dropped off and picked up by each person 
*/
Macro "School MC Postprocess"(abm)
    // Fill NSchoolDropOffs and NSchoolPickUps field in HH layer
    aggOpts.Spec = {{NSchoolDropOffs: "(SchoolForwardModeCode = 2).Count", DefaultValue: 0},
                    {NSchoolPickUps: "(SchoolReturnModeCode = 2).Count", DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)

    // Fill school return mode with forward mode for 'Bike' and 'DriveAlone' trips
    abm.CreatePersonSet({Filter: "SchoolForwardModeCode = 1 or SchoolForwardModeCode = 4", Activate: 1})
    abm.[Person.SchoolReturnMode] = abm.[Person.SchoolForwardMode]
    abm.[Person.SchoolReturnModeCode] = abm.[Person.SchoolForwardModeCode]
    abm.[Person.SchooltoHomeTime] = abm.[Person.HometoSchoolTime]

    // Update person fields to reflect # of kids dropped off or picked up (after mode choice decisions have been made)
    vwP = abm.PersonView
    
    // Dropoffs
    exprEnroute = CreateExpression(vwP, "DOEnroute", "if DropoffTourFlag = 'W1' then 1 else 0",)
    set = abm.CreatePersonSet({Filter: 'SchoolForwardModeCode = 2'})
    if set.Size > 0 then do
        aggFld = {{"DropoffPersonID", "Count",}, {exprEnroute, "Sum",}}
        vwMem = AggregateTable("Dropoffs", vwP + "|" + set.Name, "MEM", "Dropoffs", "DropoffPersonID", aggFld, null)
        {flds, specs} = GetFields(vwMem,)
        vwJ = JoinViews("PersonsAgg", GetFieldFullSpec(vwP, abm.PersonID), specs[1],)
        vecs = GetDataVectors(vwJ + "|", {flds[2], specs[3]},)
        vecsSet = null
        vecsSet.NDropOffs = nz(vecs[1])
        vecsSet.NDropOffsEnRoute = nz(vecs[2])
        SetDataVectors(vwJ + "|", vecsSet,)
        CloseView(vwJ)
        CloseView(vwMem)
    end
    DestroyExpression(GetFieldFullSpec(vwP, exprEnroute))

    // Pickups
    exprEnroute = CreateExpression(vwP, "PUEnroute", "if PickupTourFlag = 'W1' then 1 else 0",)
    set = abm.CreatePersonSet({Filter: 'SchoolReturnModeCode = 2'})
    if set.Size > 0 then do
        aggFld = {{"PickupPersonID", "Count",}, {exprEnroute, "Sum",}}
        vwMem = AggregateTable("Pickups", vwP + "|" + set.Name, "MEM", "Pickups", "PickupPersonID", aggFld, null)
        {flds, specs} = GetFields(vwMem,)
        vwJ = JoinViews("PersonsAgg", GetFieldFullSpec(vwP, abm.PersonID), specs[1],)
        vecs = GetDataVectors(vwJ + "|", {flds[2], specs[3]},)
        vecsSet = null
        vecsSet.NPickUps = nz(vecs[1])
        vecsSet.NPickUpsEnRoute = nz(vecs[2])
        SetDataVectors(vwJ + "|", vecsSet,)
        CloseView(vwJ)
        CloseView(vwMem)
    end
    DestroyExpression(GetFieldFullSpec(vwP, exprEnroute))
endMacro
