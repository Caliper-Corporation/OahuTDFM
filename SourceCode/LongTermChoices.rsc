Macro "Driver License"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    abm.[Person.License] = 2

    availSpec.Alternative = {'Yes'}
    availSpec.Expression = {'PersonHH.Age >= 16'}

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Driver License"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\DriverLicense.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "!(UnivGQStudent = 1 and Autos > 0)"})
    obj.AddUtility({UtilityFunction: Args.DriverLicenseUtility, AvailabilityExpressions: availSpec})
    obj.AddOutputSpec({ChoicesField: "License"})
    obj.ReportShares = 1
    obj.RandomSeed = 99991
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Driver License' choice model failed.")
    Args.[DriverLicense Spec] = CopyArray(ret) // For calibration purposes

    // Fill HH field with number of licensed persons
    aggOpts.Spec = {{NLicensed: "(License = 1).Count", DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)

    return(true)
endMacro


Macro "Auto Ownership"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    objT = CreateObject("Table", Args.AccessibilitiesOutputs)
    vwTAZ4Ds = objT.GetView()

    objDC = CreateObject("Table", Args.MandatoryDestAccessibility)
    vwDCAccess = objDC.GetView()

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Auto Ownership Model"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\AutoOwnership.mdl"
    obj.AddTableSource({SourceName: "HH", View: abm.HHView, IDField: abm.HHID})
    obj.AddTableSource({SourceName: "DCAccessibility", View: vwDCAccess, IDField: "TAZID"})
    obj.AddTableSource({SourceName: "TAZ4Ds", View: vwTAZ4Ds, IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "HH", Filter: "UnivGQ <> 1", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.AutoOwnershipUtility})
    obj.AddOutputSpec({ChoicesField: "Autos"})
    obj.ReportShares = 1
    obj.RandomSeed = 199999
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Auto Availability' choice model failed.")
    args.[AutoOwnership Spec] = CopyArray(ret) // For calibration purposes

    // Since the choice model returns 1 for '0 Autos', 2 for '1 Auto' etc, subtract one from the output field
    //abm.[HH.Autos] = abm.[HH.Autos] - 1
    abm.CreateHHSet({Filter: "UnivGQ <> 1", Activate: 1})
    abm.FillHHFields({Autos: "Autos - 1"}) // Alternate way

    return(true)
endMacro


Macro "Worker Models"(Args)
    RunMacro("Worker Category", Args)
    RunMacro("Work Attendance", Args)
    RunMacro("Remote Work", Args)
    Return(true)
endMacro


Macro "Worker Category"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    objT = CreateObject("Table", Args.AccessibilitiesOutputs)
    vwTAZ4Ds = objT.GetView()

    // Run Model and populate results
    modelName = "Worker Category"
    obj = CreateObject("PMEChoiceModel", {ModelName: modelName})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\WorkerCategory.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddTableSource({SourceName: "TAZ4Ds", View: vwTAZ4Ds, IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "IndustryCategory < 10 and UnivGQStudent <> 1", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.WorkerCategoryUtility})
    obj.AddOutputSpec({ChoicesField: "WorkerCategory"})
    obj.ReportShares = 1
    obj.RandomSeed = 299993
    ret = obj.Evaluate()
    if !ret then
        Throw("Running " + modelName + " choice model failed.")
    Args.(modelName + " Spec") = CopyArray(ret)

    Return(true)
endMacro


Macro "Work Attendance"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // ***** First determine number of work weekdays for full and part time residents
    workDaysSpec = Args.NumberWorkDaysSpec
    numDays = workDaysSpec.Number_Work_Weekdays
    FTShares = workDaysSpec.FullTime_Workers_Pct
    PTShares = workDaysSpec.PartTime_Workers_Pct

    if Sum(FTShares) <> 100 or Sum(PTShares) <> 100 then
        Throw("Error in 'NumberWorkDaysSpec' table. Percentage columns do not sum to 100")

    // Run over loop on FT and PT
    // Determine number of days of work and use that to determine whether person attends work on the give day
    categories = {"FT", "PT"}
    shares = {FTShares, PTShares}
    for i = 1 to categories.length do
        filter =  printf("WorkerCategory = %s and UnivGQStudent <> 1", {String(i)})
        set = abm.CreatePersonSet({Filter: filter, Activate: 1})
        if set.Size > 0 then do
            vecsSet = null
            SetRandomSeed(399989*i + 1)
            params = {population: numDays, weight: shares[i]}
            vecsSet.WorkDays = RandSamples(set.Size, "Discrete", params)

            SetRandomSeed(399989*i + 2)
            vRand = RandSamples(set.Size, "Uniform",)
            vecsSet.WorkAttendance =    if vecsSet.WorkDays = 1 and vRand <= 0.2 then 1
                                        else if vecsSet.WorkDays = 2 and vRand <= 0.4 then 1
                                        else if vecsSet.WorkDays = 3 and vRand <= 0.6 then 1
                                        else if vecsSet.WorkDays = 4 and vRand <= 0.8 then 1
                                        else if vecsSet.WorkDays = 5 then 1
                                        else 0
            abm.SetPersonVectors(vecsSet)
        end
    end
    Return(true)
endMacro


Macro "Remote Work"(Args)
    wfhPct = Args.WFHSpec.[Remote Work Percent]
    abm = RunMacro("Get ABM Manager", Args)

    // Add fields for 'WorkFromHome' and 'TravelToWork'
    newFlds = {{Name: "WorkFromHome", Type: "Short", Width: 2, Description: "Does person work from home on given day? Only filled if 'WorkAttendance = 1'"},
               {Name: "TravelToWork", Type: "Short", Width: 2, Description: "Does person travel to work on given day? 1 if WorkAttendance = 1 and WorkFromHome = 0"}}
    abm.DropPersonFields({"WorkFromHome", "TravelToWork"})
    abm.AddPersonFields(newFlds)

    // Set values for the university GQ students who are also workers or for part time workers under 18.
    set = abm.CreatePersonSet({Filter: "UnivGQStudent = 1 or Age <= 18", Activate: 1})
    v = abm.[Person.WorkAttendance]
    vecsSet = null
    vecsSet.WorkFromHome = if v = 1 then 0 else null
    vecsSet.TravelToWork = if v = 1 then 1 else null
    abm.SetPersonVectors(vecsSet)
    
    // Set values for all other workers
    set = abm.CreatePersonSet({Filter: "WorkAttendance = 1 and UnivGQStudent <> 1 and Age > 18", Activate: 1})
    SetRandomSeed(499979)
    vRand = RandSamples(set.Size, "Uniform",)

    vecs = abm.GetPersonVectors({"IndustryCategory", "WorkAttendance"})
    vInd = vecs.IndustryCategory
    arrProb = v2a(vInd).Map(do (f) Return(wfhPct[f]/100) end)
    vProb = a2v(arrProb)
    vWFH = if vRand <= vProb then 1 else 0
    
    vecsSet = null
    vecsSet.WorkFromHome = vWFH
    vecsSet.TravelToWork = if vecs.WorkAttendance = 1 and vWFH = 0 then 1 else 0
    abm.SetPersonVectors(vecsSet)
    
    Return(true)
endMacro


Macro "Mandatory Participation"(Args)
    RunMacro("Daycare Participation", Args)
    RunMacro("School Participation", Args)
    RunMacro("Univ Participation", Args)
    Return(true)
endMacro


Macro "Daycare Participation"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Daycare Participation"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\AttendDayCare.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "Age < 5", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.AttendDaycareUtility})
    obj.AddOutputSpec({ChoicesField: "AttendDayCare"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = 599999
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Daycare Participation' choice model failed.")
    Args.[Daycare Participation Spec] = CopyArray(ret)

    Return(true)
endMacro


Macro "School Participation"(Args)
    // Run Model and populate results
    attendProb = Args.SchoolStatusProb
    if attendProb < 0 or attendProb > 1 then
        Throw("Argument \'SchoolStatusProb\' should be in the range [0, 1]")

    // Choice based on attendProb value
    abm = RunMacro("Get ABM Manager", Args)
    set = abm.CreatePersonSet({Filter: "Age >= 5 and Age <= 18 and WorkerCategory <> 1", Activate: 1})
    SetRandomSeed(699967)
    v = RandSamples(set.Size, "Uniform",)
    v2 = if v <= attendProb then 1 else 0
    abm.[Person.AttendSchool] = v2

    // Fill NSchoolKids in HH table
    aggOpts.Spec = {{NSchoolKids: "(AttendSchool = 1 or AttendDaycare = 1).Count", DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)
    
    Return(true)
endMacro


Macro "Univ Participation"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "University Participation"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\AttendUniversity.mdl"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "Age >= 19 and UnivGQStudent <> 1", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.AttendUnivUtility})
    obj.AddOutputSpec({ChoicesField: "AttendUniv"})
    obj.ReportShares = Args.ReportShares
    obj.RandomSeed = 799999
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'University Participation' choice model failed.")
    Args.[Univ Participation Spec] = CopyArray(ret)

    Return(true)
endMacro


Macro "Work Location"(Args)
    // Run in a loop for the various industries
    sizeVars = Args.WorkLocSize
    sizeFlds = sizeVars.[Attractions Field]
    indCodes = sizeVars.Industry

    // Check if shadow price table already exists. If not, create an table with ID field and zero fields to store shadow prices
    ShadowPricesTable = Args.WorkDCShadowPrices
    if !GetFileInfo(ShadowPricesTable) then do
        spOpts = null
        spOpts.OutputFile = ShadowPricesTable
        spOpts.Fields = indCodes.Map(do (f) Return("Industry" + String(f)) end)
        spOpts.ReferenceData = {File: Args.DemographicOutputs, IDField: "TAZ"}
        RunMacro("Create Shadow Price Table", spOpts)
    end

    abm = RunMacro("Get ABM Manager", Args)
    pbar = CreateObject("G30 Progress Bar", "Running Work Location Model by Industry", true, indCodes.length)
    for i = 1 to indCodes.length do
        indCode = String(indCodes[i])
        
        utilSpec = null
        utilSpec.UtilityFunction = Args.WorkLocUtility
        utilSpec.SubstituteStrings = {{"<IndCode>", indCode}}
        
        // Find BG Location
        obj = CreateObject("PMEChoiceModel", {ModelName: "Work Location: Industry " + indCode})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\WorkLocation_Ind" + indCode + ".dcm"
        obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "WorkDCShadowPrices", File: ShadowPricesTable, IDField: "TAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.MandatoryModeAccessibility, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "PersonHH", Filter: "WorkIndustry = " + indCode, OField: "TAZID"})
        obj.AddUtility(utilSpec)
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: sizeFlds[i]})
        if Args.WorkSPFlag then do
            tempOutputSP = GetTempPath() + "ShadowPrice_Industry" + indCode + ".bin"
            obj.AddShadowPrice({TargetName: "TAZData", TargetField: sizeFlds[i], 
                                Iterations: 10, Tolerance: 0.01, OutputShadowPriceTable: tempOutputSP})
        end
        obj.AddOutputSpec({ChoicesField: "WorkTAZ"})
        obj.RandomSeed = 899981 + 42*i
        ret = obj.Evaluate()
        if !ret then
            Throw("Running 'Work Location' choice model for Industry " + indCode + " failed.")
        obj = null
        
        // Copy values from temporary shadow price table to main shadow price table
        if Args.WorkSPFlag then do
            spOpts = null
            spOpts.SourceFile = tempOutputSP
            spOpts.TargetData = {File: ShadowPricesTable, IDField: "TAZ", OutputField: "Industry" + indCode}
            RunMacro("Copy Shadow Prices", spOpts)
        end

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Fill home to work times
    spec = {abmManager: abm,
            OField: "TAZID",
            DField: "WorkTAZ",
            FillField: "HometoWorkTime",
            Filter: "WorkIndustry <= 10",
            Matrix: {Name: Args.HighwaySkimAM, Core: "Time"}}
    RunMacro("Fill from matrix", spec)

    Return(true)
endMacro


Macro "Univ Location"(Args)
    // Create temporary University Enrollment after subtracting dorm population that have been acocunnted for
    enrollmentFld = RunMacro("Adjust University Enrollment", Args.DemographicOutputs)

    ShadowPricesTable = Args.UnivDCShadowPrices
    if !GetFileInfo(ShadowPricesTable) then do
        spOpts = null
        spOpts.OutputFile = ShadowPricesTable
        spOpts.Fields = {"ShadowPrice"}
        spOpts.ReferenceData = {File: Args.DemographicOutputs, IDField: "TAZ"}
        RunMacro("Create Shadow Price Table", spOpts)
    end

    // Run Model and populate results
    abm = RunMacro("Get ABM Manager", Args)
    obj = CreateObject("PMEChoiceModel", {ModelName: "University TAZ Location"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\UniversityLocation.dcm"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
    obj.AddTableSource({SourceName: "UnivDCShadowPrices", File: ShadowPricesTable, IDField: "TAZ"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "AttendUniv = 1 and UnivGQStudent <> 1", OField: "TAZID"})
    obj.AddUtility({UtilityFunction: Args.UnivLocUtility})
    obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
    obj.AddSizeVariable({Name: "TAZData", Field: enrollmentFld})
    if Args.UnivSPFlag then do // Perform shadow pricing only if shadow price table does not already exist
        tempOutputSP = GetTempPath() + "ShadowPrice_University.bin"
        obj.AddShadowPrice({TargetName: "TAZData", TargetField: enrollmentFld, 
                            Iterations: 10, Tolerance: 0.01, OutputShadowPriceTable: tempOutputSP})
    end
    obj.AddOutputSpec({ChoicesField: "UnivTAZ"})
    obj.RandomSeed = 999983
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'University Location' choice model failed.")

    // Copy values from temporary shadow price table to main shadow price table
    if Args.UnivSPFlag then do
        spOpts = null
        spOpts.SourceFile = tempOutputSP
        spOpts.TargetData = {File: ShadowPricesTable, IDField: "TAZ", OutputField: "ShadowPrice"}
        RunMacro("Copy Shadow Prices", spOpts)
    end

    // Fill home to univ times
    spec = {abmManager: abm,
            OField: "TAZID",
            DField: "UnivTAZ",
            FillField: "HometoUnivTime",
            Filter: "AttendUniv = 1",
            Matrix: {Name: Args.HighwaySkimAM, Core: "Time"}}
    RunMacro("Fill from matrix", spec)

    Return(true)
endMacro


// Create enrollment field for univerity DC by subtracting out dorm students enrollment
Macro "Adjust University Enrollment" (TAZData)
    // Open TAZ and export relevant information into temp file
    objT = CreateObject("Table", TAZData)
    n = objT.SelectByQuery({SetName: "Selection", Query: "UnivGQ = 1"})
    vwT = objT.GetView()
    
    if n = 0 then
        outField = 'UnivEnrollment'
    else do
        outField = 'UnivEnrollmentforDC'
        newFlds = newFlds + {{FieldName: outField, Type: "real", Width: 12, Decimals: 2}}
        objT.AddFields({Fields: newFlds})

        // Aggregate univ residents by univ zone
        aggSpec = {{"GroupQuarterPopulation", "Sum",}}
        vwAgg = AggregateTable("MemAggr", vwT + "|Selection", "MEM",, "UnivTAZ", aggSpec,)
        {flds, specs} = GetFields(vwAgg,)

        // Fill new field
        vwJ = JoinViews("TAZAggr", GetFieldFullSpec(vwT, "TAZ"), specs[1],)
        vecs = GetDataVectors(vwJ + "|", {"UnivEnrollment", specs[2]},)
        v = if nz(vecs[2]) > vecs[1] then null else vecs[1] - nz(vecs[2])
        SetDataVector(vwJ + "|", outField, nz(v), )
        CloseView(vwJ)
        CloseView(vwAgg)
    end
    objT = null
    Return(outField)
endMacro


// School Location Choice
Macro "School Location"(Args)
    // Define school segments
    filters = {"Age >= 5 and Age <= 11", "Age >= 12 and Age <= 14", "Age >= 15 and Age <= 18"}
    categories = {"Elementary", "Middle", "HighSchool"}

    availExpressions = null
    availExpressions.Alternative = {"Destinations"}
    availExpressions.Expression = {"AutoSkim.Time <= 60"}

    // Check if shadow price table already exists. If not, create an table with ID field and zero fields to store shadow prices
    ShadowPricesTable = Args.SchoolDCShadowPrices
    if !GetFileInfo(ShadowPricesTable) then do
        spOpts = null
        spOpts.OutputFile = ShadowPricesTable
        spOpts.Fields = CopyArray(categories)
        spOpts.ReferenceData = {File: Args.DemographicOutputs, IDField: "TAZ"}
        RunMacro("Create Shadow Price Table", spOpts)
    end
    
    abm = RunMacro("Get ABM Manager", Args)
    pbar = CreateObject("G30 Progress Bar", "Running School Location Model by grade", true, filters.length)
    utilSpec = null
    utilSpec.UtilityFunction = Args.SchoolLocUtility
    utilSpec.AvailabilityExpressions = availExpressions
    
    for i = 1 to categories.length do
        type = categories[i]
        utilSpec.SubstituteStrings = {{"<Type>", type}}
        
        // Run Model and populate results
        obj = CreateObject("PMEChoiceModel", {ModelName: "School Location: " + type})
        obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\SchoolLocation_" + type + ".dcm"
        obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
        obj.AddTableSource({SourceName: "SchoolDCShadowPrices", File: ShadowPricesTable, IDField: "TAZ"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddPrimarySpec({Name: "PersonHH", Filter: filters[i], OField: "TAZID"})
        obj.AddUtility(utilSpec)
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: type + "Enrollment"})
        if Args.SchoolSPFlag then do // Perform shadow pricing only if shadow price table does not already exist
            tempOutputSP = GetTempPath() + "ShadowPrice_" + type + ".bin"
            obj.AddShadowPrice({TargetName: "TAZData", TargetField: type + "Enrollment", 
                                Iterations: 10, Tolerance: 0.01, OutputShadowPriceTable: tempOutputSP})
        end
        obj.AddOutputSpec({ChoicesField: "SchoolTAZ"})
        obj.RandomSeed = 1099997
        ret = obj.Evaluate()
        if !ret then
            Throw("Running 'School Location' choice model for " + type + " failed")
        
        // Copy values from temporary shadow price table to main shadow price table
        if Args.SchoolSPFlag then do
            spOpts = null
            spOpts.SourceFile = tempOutputSP
            spOpts.TargetData = {File: ShadowPricesTable, IDField: "TAZ", OutputField: type}
            RunMacro("Copy Shadow Prices", spOpts)
        end
        
        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Fill home to univ times
    spec = {abmManager: abm,
            OField: "TAZID",
            DField: "SchoolTAZ",
            FillField: "HometoSchoolTime",
            Filter: "AttendSchool = 1",
            Matrix: {Name: Args.HighwaySkimAM, Core: "Time"}}
    RunMacro("Fill from matrix", spec)

    Return(true)
endMacro


// School Location Choice
Macro "Daycare Location"(Args)
    availExpressions = null
    availExpressions.Alternative = {"Destinations"}
    availExpressions.Expression = {"AutoSkim.Distance <= 4"}

    abm = RunMacro("Get ABM Manager", Args)
    
    utilSpec = null
    utilSpec.UtilityFunction = Args.DaycareLocUtility
    utilSpec.AvailabilityExpressions = availExpressions

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Daycare Location"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\DaycareLocation.dcm"
    obj.AddTableSource({SourceName: "PersonHH", View: abm.PersonHHView, IDField: abm.PersonID})
    obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "PersonHH", Filter: "AttendDaycare = 1", OField: "TAZID"})
    obj.AddUtility(utilSpec)
    obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
    obj.AddOutputSpec({ChoicesField: "SchoolTAZ"})
    obj.RandomSeed = 1199999
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Daycare Location' choice model failed")
        
    // Fill home to univ times
    spec = {abmManager: abm,
            OField: "TAZID",
            DField: "SchoolTAZ",
            FillField: "HometoSchoolTime",
            Filter: "AttendDaycare = 1",
            Matrix: {Name: Args.HighwaySkimAM, Core: "Time"}}
    RunMacro("Fill from matrix", spec)

    Return(true)
endMacro




// Create empty shadow price table with required structure and fill shadow prices with zeros
Macro "Create Shadow Price Table"(spOpts)
    obj = CreateObject("Table", spOpts.ReferenceData.File)
    idFld = spOpts.ReferenceData.IDField
    objSP = obj.Export({FileName: spOpts.OutputFile, FieldNames: {idFld}})

    for fld in spOpts.Fields do
        flds = flds + {{FieldName: fld, Type: "Real"}}
    end
    objSP.AddFields({Fields: flds})

    v = objSP.(idFld)
    vZero = Vector(v.length, "Double", {{"Constant", 0.0}})
    
    vecsSet = null
    for fld in spOpts.Fields do
        vecsSet.(fld) = vZero
    end
    objSP.SetDataVectors({FieldData: vecsSet})

    objSP = null
    obj = null
endMacro


// Copy shadow price values from temporary output table to the shadow prices table
Macro "Copy Shadow Prices"(spOpts)
    objSP = CreateObject("Table", spOpts.TargetData.File)
    objT = CreateObject("Table", spOpts.SourceFile)
    objJ = objT.Join({Table: objSP, LeftFields: {"ID"}, RightFields: {spOpts.TargetData.IDField}})
    outfld = spOpts.TargetData.OutputField
    objJ.(outfld) = objJ.(outfld) + objJ.Shadow
    objJ = null
    objT = null
    objSP = null
endMacro
