/*
    - Generate MC accessibiltiies (modified logsums) by mode group using an aggregate mode choicep specification
      Generate accessibility matrices by auto, non motorized and transit mode groups

    - Generate DC accessibilities (modified logsums) by origin zone using an aggregate DC specification that uses the mode choice logsums
*/
Macro "Mandatory Accessibility"(Args)
    // Generate MC accessibilities from the aggregate mode choice spec
    RunMacro("Mand MC Accessibility", Args)

    // Generate DC accessibilities from the aggregate destination choice spec
    RunMacro("Mand DC Accessibility", Args)

    Return(true)
endMacro


/*
    Given the aggregate mode choice spec, generate accessibility matrices by mode group

    - For a given mode group, disable modes not part of the group, run model and obtain the logsum matrix
    - Compute the accessibilty core as ln(1 + exp(Root_Logsum))
*/
Macro "Mand MC Accessibility"(Args)
    // Create empty output MC logsum matrix
    mSpec = {TAZFile: Args.TAZGeography, OutputFile: Args.MandatoryModeAccessibility, DataType: "Double", 
             Cores: {"Auto", "NM", "PT"}, Label: "Mandatory Mode Accessibility"}
    mat = RunMacro("Create Empty Matrix", mSpec)
    mOutObj = CreateObject("Matrix", mat)

    // Define mode groups
    modeGroups.Auto = {"DriveAlone", "Carpool", "Other"}
    modeGroups.NM = {"Bike", "Walk"}
    modeGroups.PT = {"Bus"}
    allModes = modeGroups.Auto + modeGroups.NM + modeGroups.PT

    for modeGroup in modeGroups do
        mainMode = modeGroup[1]
        subModes = modeGroup[2]
        
        // Get availability parameter and modify it
        inputAvail = Args.MandatoryAggMCAvail
        avail = CopyArray(inputAvail)
        availAlts = avail.Alternative // Array of alternatives in current availability specification
        for mode in allModes do
            pos = subModes.position(mode)
            if pos > 0 then 
                continue

            altPos = availAlts.position(mode)
            if altPos > 0 then  // replace avail
                avail.Expression[altPos] = "TAZ4Ds.TAZID.O < 0"    
            else do             // add new avail term
                avail.Alternative = avail.Alternative + {mode}
                avail.Expression = avail.Expression + {"TAZ4Ds.TAZID.O < 0"}  // Always unavailable
            end
        end

        // Run PME model
        util = null
        util.UtilityFunction = Args.MandatoryAggMCUtility
        util.SubstituteStrings = {{"<VOT>", String(Args.VOT)}, {"<AOC>", String(Args.AOC)}}
        util.AvailabilityExpressions = avail

        type = mainMode + "AggregateMC"
        logsumFile = printf("%s\\Intermediate\\MCLogsum_%s.mtx", {Args.[Output Folder], type})
        probFile = printf("%s\\Intermediate\\MCProb_%s.mtx", {Args.[Output Folder], type})
        utilFile = printf("%s\\Intermediate\\MCUtil_%s.mtx", {Args.[Output Folder], type})
        
        obj = CreateObject("PMEChoiceModel", {ModelName: type})
        obj.OutputModelFile = printf("%s\\Intermediate\\%s.mdl", {Args.[Output Folder], type})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "PTSkim", File: Args.TransitWalkSkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddPrimarySpec({Name: "AutoSkim"})
        obj.AddUtility(util)
        obj.AddAlternatives({AlternativesTree: Args.MandatoryAggMCModes})
        obj.AddOutputSpec({Probability: probFile, Logsum: logsumFile, Utility: utilFile})
        ret = obj.Evaluate()

        // Get logsums, write into new matrix
        mObj = CreateObject("Matrix", logsumFile)
        mOutObj.(mainMode) := log(1 + nz(exp(mObj.Root)))
        mObj = null
    end
    mOutObj = null
    mat = null
endMacro


/*
    Generate DC accessibility vectors by mode type (Auto, NM, PT)
    Each vector represents destinations accessible by that mode
    - Compute the accessibilty vector as the log as marginal row sum of the exponentiated utility matrix
*/
Macro "Mand DC Accessibility"(Args)
    // Create empty output table
    objTAZ = CreateObject("Table", Args.TAZGeography)
    objOut = objTAZ.Export({FileName: Args.MandatoryDestAccessibility, FieldNames: {"TAZID"}})
    fldsAdd = {{FieldName: "DestAccessibilityAuto", Type: "real", Width: 12, Decimals: 2},
                {FieldName: "DestAccessibilityNM", Type: "real", Width: 12, Decimals: 2},
                {FieldName: "DestAccessibilityPT", Type: "real", Width: 12, Decimals: 2}}
    objOut.AddFields({Fields: fldsAdd})
    sortOrder = {{"TAZID", "Ascending"}}
    objOut.Sort({FieldArray: sortOrder})
    objTAZ = null

    // Run DC model for each mode, and fill the column in the table
    modes = {"Auto", "NM", "PT"}
    for mode in modes do
        tag = "DC_" + mode

        // Compute Size Variable and fill field in output TAZ demographics table
        sizeFld = tag + "_SizeVar"
        opt = null
        opt.TableObject = CreateObject("Table", Args.DemographicOutputs)
        opt.Equation = Args.(sizeFld)   // e.g. Args.DC_Auto_SizeVar
        opt.NewOutputField = sizeFld
        RunMacro("Compute Size Variable", opt)
        opt.TableObject = null

        // Run DC Model
        tmpUtilFile = printf("%s\\Intermediate\\%s_Util.mtx", {Args.[Output Folder], tag})
        obj = CreateObject("PMEChoiceModel", {ModelName: tag})
        obj.OutputModelFile = printf("%s\\Intermediate\\%s.dcm", {Args.[Output Folder], tag})
        obj.AddTableSource({SourceName: "TAZ4Ds", File: Args.AccessibilitiesOutputs, IDField: "TAZID"})
        obj.AddTableSource({SourceName: "TAZData", File: Args.DemographicOutputs, IDField: "TAZ"})
        obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
        obj.AddMatrixSource({SourceName: "Intrazonal", File: Args.IZMatrix, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddMatrixSource({SourceName: "ModeAccessibility", File: Args.MandatoryModeAccessibility, RowIndex: "TAZ", ColIndex: "TAZ"})
        obj.AddPrimarySpec({Name: "AutoSkim"})
        obj.AddUtility({UtilityFunction: Args.(tag + "_Utility")})
        obj.AddDestinations({DestinationsSource: "AutoSkim", DestinationsIndex: "InternalTAZ"})
        obj.AddSizeVariable({Name: "TAZData", Field: sizeFld})
        obj.AddOutputSpec({Probability: GetTempPath() + "Prob.mtx", Utility: tmpUtilFile})
        ret = obj.Evaluate()
        if !ret then
            Throw("Model Run failed while computing TAZ destination logsums for " + tag)

        // Open Utility Matrix, add exp(u) core, get row marginal sums, apply ln() and write to output table
        mObj = CreateObject("Matrix", tmpUtilFile)
        mObj.AddCores({"ExpUtil"})
        coreNames = mObj.GetCoreNames()
        mObj.ExpUtil := exp(mObj.(coreNames[1]))
        vLS = mObj.GetVector({Core: "ExpUtil", Marginal: "Row Sum"})
        vLS = if vLS = null then null else log(vLS)
        mObj = null

        outFld = "DestAccessibility" + mode
        objOut.(outFld) = vLS
    end
    objOut = null
endMacro
