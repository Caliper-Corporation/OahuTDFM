Macro "Work Tours Frequency"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    // Run Model for workers and populate results
    obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: "Work Tours Frequency"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\Work_MandatoryTours.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "TravelToWork = 1", OField: "TAZID", DField: "WorkTAZ"})
    obj.AddUtility({UtilityFunction: Args.WorkMandatoryFreqUtility})
    obj.AddOutputSpec({ChoicesField: "NumberWorkTours"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = 1299989
    ret = obj.Evaluate()
    if !ret then
        Throw("Running Work Tours Frequency model failed.")
    Args.[Work Tours Frequency Spec] = CopyArray(ret)
    obj = null
    Return(true)
endMacro


Macro "Univ Tours Frequency"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    // Run Model for univ students and populate results
    obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: "University Tours Frequency"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\Univ_MandatoryTours.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "AttendUniv = 1", OField: "TAZID", DField: "UnivTAZ"})
    obj.AddUtility({UtilityFunction: Args.UnivMandatoryFreqUtility})
    obj.AddOutputSpec({ChoicesField: "NumberUnivTours"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = 1399999
    ret = obj.Evaluate()
    if !ret then
        Throw("Running University Tours Frequency model failed.")
    Args.[Univ Tours Frequency Spec] = CopyArray(ret)
    obj = null
    Return(true)
endMacro


Macro "FullTime Work Act Dur"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-20":   0.4,
                        "20-40":  0.3,
                        "40-50":  0.2,
                        "50-60":  0.1}

    // First Tour
    Opts = {abmManager: abm,
            ModelName: "Full Time Workers: Work Tour 1 Duration",
            ModelFile: "FTWorkers_MandatoryDuration.mdl",
            Filter: "WorkerCategory = 1 and WorkAttendance = 1", // Include WFH workers here
            DestField: "WorkTAZ", 
            Alternatives: Args.FTWorkDurAlts,
            Utility: Args.FTWorkDurUtility,
            ChoiceField: "Work1_DurChoice",
            SimulatedTimeField: "Work1_Duration",
            HourlyProfile: hrlyProfileArray,
            MinimumDuration: 30,
            RandomSeed: 1499977}
    RunMacro("Mandatory Activity Time", Args, Opts)

    // Second Tour
    Opts = {abmManager: abm,
            ModelName: "Full Time Workers: Work Tour 2 Duration",
            ModelFile: "FTWorkers_MandatoryDuration2.mdl",
            Filter: "WorkerCategory = 1 and NumberWorkTours = 2",   // No need to consider WFH here
            DestField: "WorkTAZ", 
            Alternatives: Args.FTWorkDurAlts,
            Utility: Args.FTWorkDurUtility,
            ChoiceField: "Work2_DurChoice",
            SimulatedTimeField: "Work2_Duration",
            HourlyProfile: hrlyProfileArray,
            MinimumDuration: 30,
            RandomSeed: 1599977}
    RunMacro("Mandatory Activity Time", Args, Opts)
    Return(true)
endMacro


Macro "University Act Dur"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-20":   0.4,
                        "20-40":  0.3,
                        "40-50":  0.2,
                        "50-60":  0.1}

    // First Tour
    Opts = {abmManager: abm,
            ModelName: "University Tour 1 Duration",
            ModelFile: "Univ_MandatoryDuration.mdl",
            Filter: "AttendUniv = 1",
            DestField: "UnivTAZ", 
            Alternatives: Args.UnivDurAlts,
            Utility: Args.UnivDurUtility,
            ChoiceField: "Univ1_DurChoice",
            SimulatedTimeField: "Univ1_Duration",
            HourlyProfile: hrlyProfileArray,
            MinimumDuration: 30,
            RandomSeed: 1699993} 
    RunMacro("Mandatory Activity Time", Args, Opts)

    // Second Tour
    Opts = {abmManager: abm,
            ModelName: "University Tour 2 Duration",
            ModelFile: "Univ_MandatoryDuration2.mdl",
            Filter: "AttendUniv = 1 and NumberUnivTours = 2",
            DestField: "UnivTAZ", 
            Alternatives: Args.UnivDurAlts,
            Utility: Args.UnivDurUtility,
            ChoiceField: "Univ2_DurChoice",
            SimulatedTimeField: "Univ2_Duration",
            HourlyProfile: hrlyProfileArray,
            MinimumDuration: 30,
            RandomSeed: 1799999} 
    RunMacro("Mandatory Activity Time", Args, Opts)
    Return(true)
endMacro


Macro "PartTime Work Act Dur"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-10":   0.15,
                        "10-20":  0.25,
                        "20-40":  0.3,
                        "40-50":  0.2,
                        "50-60":  0.1}

    // First Tour
    Opts = {abmManager: abm,
            ModelName: "Part Time Workers: Work Tour 1 Duration",
            ModelFile: "PTWorkers_MandatoryDuration.mdl",
            Filter: "WorkerCategory = 2 and WorkAttendance = 1",    // Note, WFH workers will have a duration
            DestField: "WorkTAZ", 
            Alternatives: Args.PTWorkDurAlts,
            Utility: Args.PTWorkDurUtility,
            ChoiceField: "Work1_DurChoice",
            SimulatedTimeField: "Work1_Duration",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 1899983} 
    RunMacro("Mandatory Activity Time", Args, Opts)

    // Second Tour
    Opts = {abmManager: abm,
            ModelName: "Part Time Workers: Work Tour 2 Duration",
            ModelFile: "PTWorkers_MandatoryDuration.mdl",
            Filter: "WorkerCategory = 2 and NumberWorkTours = 2", // No need to consider WFH here
            DestField: "WorkTAZ", 
            Alternatives: Args.PTWorkDurAlts,
            Utility: Args.PTWorkDurUtility,
            ChoiceField: "Work2_DurChoice",
            SimulatedTimeField: "Work2_Duration",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 1999993} 
    RunMacro("Mandatory Activity Time", Args, Opts)
    Return(true)
endMacro


Macro "School Act Dur"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    hrlyProfileArray = {"0-10":   0.1,
                        "10-20":  0.25,
                        "20-40":  0.3,
                        "40-50":  0.2,
                        "50-60":  0.15}

    Opts = {abmManager: abm,
            ModelName: "School Tour Duration",
            ModelFile: "School_Duration.mdl",
            Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
            DestField: "SchoolTAZ", 
            Alternatives: Args.SchDurAlts,
            Utility: Args.SchDurUtility,
            ChoiceField: "School_DurChoice",
            SimulatedTimeField: "School_Duration",
            HourlyProfile: hrlyProfileArray,
            MinimumDuration: 30,
            RandomSeed: 2099963} 
    RunMacro("Mandatory Activity Time", Args, Opts)
    Return(true)
endMacro


/*
    Activity start time models
*/
Macro "FullTime Work Start"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-10":   0.4,
                        "10-40":  0.3,
                        "40-50":  0.15,
                        "50-60":  0.15}

    // Run Start Model for first tour
    Opts = {abmManager: abm,
            ModelName: "Full Time Workers: Work Tour 1 Start Time",
            ModelTag: "Work1",
            ModelFile: "FTWorkers_StartTime1.mdl",
            Filter: "WorkerCategory = 1 and WorkAttendance = 1", // This includes WFH
            DestField: "WorkTAZ",
            Alternatives: Args.FTWorkStartAlts,
            Utility: Args.FTWorkStartUtility,
            ChoiceField: "Work1_StartInt",
            SimulatedTimeField: "Work1_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2199979} 
    Args.[FullTime Work Tour1 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    
    // Determine availabilities for second tour. (Model alternatives are constrained by first tour choices)
    spec = {FirstActivityStart: "Work1_StartTime", 
            FirstActivityDuration: "Work1_Duration",
            TravelTime: "HometoWorkTime",
            MinTourSpacing: 45,
            Alternatives: Args.FTWorkStartAlts}
    availArr = RunMacro("Get StartInt Avail", spec)

    // Run Start Model for second tour
    Opts = {abmManager: abm,
            ModelName: "Full Time Workers: Work Tour 2 Start Time",
            ModelTag: "Work2",
            ModelFile: "FTWorkers_StartTime2.mdl",
            Filter: "WorkerCategory = 1 and NumberWorkTours = 2",
            DestField: "WorkTAZ",
            Alternatives: Args.FTWorkStartAlts,
            Availabilities: availArr, // Additional input containing availabilities array
            Utility: Args.FTWorkStartUtility,
            ChoiceField: "Work2_StartInt",
            SimulatedTimeField: "Work2_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2299963} 
    Args.[FullTime Work Tour2 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    Return(true)
endMacro


Macro "University Start"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-10":   0.4,
                        "10-40":  0.3,
                        "40-50":  0.15,
                        "50-60":  0.15}

    // 1st University tour for persons who do not work
    Opts = {abmManager: abm,
            ModelName: "University Tour Start Hour",
            ModelTag: "Univ1",
            ModelFile: "Univ_StartTime1.mdl",
            Filter: "AttendUniv = 1 and WorkAttendance <> 1",
            DestField: "UnivTAZ",
            Alternatives: Args.UnivStartAlts,
            Utility: Args.UnivStartUtility,
            ChoiceField: "Univ1_StartInt",
            SimulatedTimeField: "Univ1_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2399993} 
    Args.[Univ Tour1 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    
    // Determine availabilities for second tour. (Model alternatives are constrained by first tour i.e. "Univ1" choices)
    spec = {FirstActivityStart: "Univ1_StartTime", 
            FirstActivityDuration: "Univ1_Duration",
            TravelTime: "HometoUnivTime",
            MinTourSpacing: 45,
            Alternatives: Args.UnivStartAlts}
    availArr = RunMacro("Get StartInt Avail", spec)

    // Run Start Model for second tour
    Opts = {abmManager: abm,
            ModelName: "University Tour Start Hour",
            ModelTag: "Univ2",
            ModelFile: "Univ_StartTime2.mdl",
            Filter: "AttendUniv = 1 and NumberUnivTours > 1",
            DestField: "UnivTAZ",
            Alternatives: Args.UnivStartAlts,
            Availabilities: availArr, // Additional input containing availabilities array
            Utility: Args.UnivStartUtility,
            ChoiceField: "Univ2_StartInt",
            SimulatedTimeField: "Univ2_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2499997} 
    Args.[Univ Tour2 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    // Determine availabilities for first univ tour for people who also go to work. Scheduled after work.
    spec = {FirstActivityStart: "Work1_StartTime", 
            FirstActivityDuration: "Work1_Duration",
            TravelTime: "HometoUnivTime",
            MinTourSpacing: 45,
            Alternatives: Args.UnivStartAlts}
    availArr = RunMacro("Get StartInt Avail", spec)

    // Run Start Model Ist university tour for persons who also work
    Opts = {abmManager: abm,
            ModelName: "University Tour Start Hour",
            ModelTag: "Univ1",
            ModelFile: "Univ_StartTime1_Wrk.mdl",
            Filter: "AttendUniv = 1 and WorkAttendance = 1",
            DestField: "UnivTAZ",
            ModeField: "UnivMode", 
            Alternatives: Args.UnivStartAlts,
            Availabilities: availArr, // Additional input containing availabilities array
            Utility: Args.UnivStartUtility,
            ChoiceField: "Univ1_StartInt",
            SimulatedTimeField: "Univ1_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2599999} 
    Args.[Univ Tour2 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)
    
    Return(true)
endMacro


Macro "PartTime Work Start"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-10":   0.30,
                        "10-20":  0.20,
                        "20-30":  0.10,
                        "30-40":  0.15,
                        "40-50":  0.15,
                        "50-60":  0.10}

    // Run Start Model for first tour
    Opts = {abmManager: abm,
            ModelName: "Part Time Workers: Work Tour 1 Start Time",
            ModelTag: "Work1",
            ModelFile: "PTWorkers_StartTime1.mdl",
            Filter: "WorkerCategory = 2 and WorkAttendance = 1 and AttendSchool <> 1",
            DestField: "WorkTAZ",
            Alternatives: Args.PTWorkStartAlts,
            Utility: Args.PTWorkStartUtility,
            ChoiceField: "Work1_StartInt",
            SimulatedTimeField: "Work1_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2699999} 
    Args.[PartTime Work Tour1 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    
    // Determine availabilities for second tour. (Model alternatives are constrained by first tour choices)
    spec = {FirstActivityStart: "Work1_StartTime", 
            FirstActivityDuration: "Work1_Duration",
            TravelTime: "HometoWorkTime",
            MinTourSpacing: 45,
            Alternatives: Args.PTWorkStartAlts}
    availArr = RunMacro("Get StartInt Avail", spec)

    // Run Start Model for second tour
    Opts = {abmManager: abm,
            ModelName: "Part Time Workers: Work Tour 2 Start Time",
            ModelTag: "Work2",
            ModelFile: "PTWorkers_StartTime2.mdl",
            Filter: "WorkerCategory = 2 and NumberWorkTours = 2 and AttendSchool <> 1",
            DestField: "WorkTAZ",
            Alternatives: Args.PTWorkStartAlts,
            Availabilities: availArr, // Additional input containing availabilities array
            Utility: Args.PTWorkStartUtility,
            ChoiceField: "Work2_StartInt",
            SimulatedTimeField: "Work2_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2799991}
    Args.[PartTime Work Tour2 Start Spec] = RunMacro("Mandatory Activity Time", Args, Opts)
    
    Return(true)
endMacro


Macro "School Start"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    hrlyProfileArray = {"0-10":   0.15,
                        "10-20":  0.20,
                        "20-40":  0.25,
                        "40-60":  0.40}

    Opts = {abmManager: abm,
            ModelName: "School Tour Start Hour",
            ModelTag: "School",
            ModelFile: "School_StartTime.mdl",
            Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
            DestField: "SchoolTAZ", 
            Utility: Args.SchStartUtility,
            ChoiceField: "School_StartInt",
            SimulatedTimeField: "School_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2899997} 
    Args.[School Tour Dep Spec] = RunMacro("Mandatory Activity Time", Args, Opts)

    
    // Determine availabilities for second tour. (Model alternatives are constrained by first tour choices)
    // Note "School" tour imposes constraints on part time workers who also attend school. School tour scheduled first.
    spec = {FirstActivityStart: "School_StartTime", 
            FirstActivityDuration: "School_Duration",
            TravelTime: "HometoWorkTime",
            MinTourSpacing: 45,
            Alternatives: Args.PTWorkStartAlts}
    availArr = RunMacro("Get StartInt Avail", spec)

    // Run Start Model for first part time work tour after completion of school tour
    Opts = {abmManager: abm,
            ModelName: "PTWorkers who attend School: Work Tour 1 Start Time",
            ModelTag: "Work1",
            ModelFile: "PTWorkersSch_StartTime.mdl",
            Filter: "WorkerCategory = 2 and WorkAttendance = 1 and AttendSchool = 1",
            DestField: "WorkTAZ",
            Alternatives: Args.PTWorkStartAlts,
            Availabilities: availArr, // Additional input containing availabilities array
            Utility: Args.PTWorkStartUtility,
            ChoiceField: "Work1_StartInt",
            SimulatedTimeField: "Work1_StartTime",
            HourlyProfile: hrlyProfileArray,
            RandomSeed: 2999999}
    Args.[PT Work1 Tour Dep Spec] = RunMacro("Mandatory Activity Time", Args, Opts)
    
    Return(true)
endMacro


// Macro that runs the activity time choice model to generate activity duration.
// Called for work tour, univ tour or school tour choices
Macro "Mandatory Activity Time"(Args, Opts)
    abm = Opts.abmManager
    filter = Opts.Filter

    // Basic Check
    if Opts.Utility = null or Opts.ModelName = null or Opts.ModelFile = null or filter = null
        or Opts.DestField or Opts.ChoiceField = null then
            Throw("Invalid inputs to macro 'Activity Time Choice'")

    // Get Utility Options
    utilOpts = null
    utilOpts.UtilityFunction = Opts.Utility
    if Opts.SubstituteStrings <> null then
        utilOpts.SubstituteStrings = Opts.SubstituteStrings
    if Opts.Availabilities <> null then
        utilOpts.AvailabilityExpressions = Opts.Availabilities
    
    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {SourcesObject: Args.SourcesObject, ModelName: Opts.ModelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\" + Opts.ModelFile
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddMatrixSource({SourceName: "AMAutoSkim", File: Args.HighwaySkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddMatrixSource({SourceName: "OPAutoSkim", File: Args.HighwaySkim, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: filter, OField: "TAZID", DField: Opts.DestField})
    if Opts.Alternatives <> null then
        obj.AddAlternatives({AlternativesList: Opts.Alternatives})
    obj.AddUtility(utilOpts)
    obj.AddOutputSpec({ChoicesField: Opts.ChoiceField})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = Opts.RandomSeed
    ret = obj.Evaluate()
    if !ret then
        Throw("Model 'Mandatory Activity Choice' failed for " + Opts.ModelName)
    Args.(Opts.ModelName +  " Spec") = CopyArray(ret)

    vw = abm.PersonView
    // Simulate time after choice of interval is made
    simFld = Opts.SimulatedTimeField
    if simFld <> null then do
        set = abm.CreatePersonSet({Filter: filter, Activate: 1})
        if set.Size > 0 then do
            // Simulate duration in minutes for duration choice predicted above
            opt = null
            opt.ViewSet = vw + "|" + set.Name
            opt.InputField = Opts.ChoiceField
            opt.OutputField = simFld
            opt.HourlyProfile = Opts.HourlyProfile
            opt.AlternativeIntervalInMin = Opts.AlternativeIntervalInMin
            RunMacro("Simulate Time", opt)
        end
    end

    // Set minimum duration if specified
    minDur = Opts.MinimumDuration
    if minDur > 0 then do
        qry = printf("(%s) and (%s < %s)", {filter, Opts.SimulatedTimeField, string(minDur)})
        set = abm.CreatePersonSet({Filter: qry, Activate: 1})
        if set.Size > 0 then do
            v = Vector(set.Size, "Long", {{"Constant", minDur}})
            vecsSet = null
            vecsSet.(simFld) = v
            abm.SetPersonVectors(vecsSet)
        end
    end
endMacro


// Determines availability expressions for each of the start time alternatives.
// Approximates earliest time a person can start the next activity
Macro "Get StartInt Avail"(spec)
    altSpec = spec.Alternatives
    alts = altSpec.Alternative
    availArr = null
    availArr.Alternative = CopyArray(alts)

    for alt in alts do
        tmpArr = ParseString(alt, "- ")
        altStart = s2i(tmpArr[1]) * 60 // Minutes from midnight
        expr = printf("if PersonHH.%s + PersonHH.%s + 2*PersonHH.%s + %s < %s then 1 else 0", 
                      {spec.FirstActivityStart, spec.FirstActivityDuration, spec.TravelTime, String(spec.MinTourSpacing), String(altStart)})
        availArr.Expression =  availArr.Expression + {expr}
    end

    Return(availArr)
endMacro
