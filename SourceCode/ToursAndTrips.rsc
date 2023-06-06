/************************** Tour file creation macros ****************************************
/* 
    Create master tour file
    1. Create tours from person decisions
    2. Modify tours based on PUDO stops
    3. Create separate Dropoff/Pickup tours table as necessary
    4. Concatenate adjusted tours with dropoff and pickup tours tables
*/
Macro "Create Mandatory Tour File"(Args)
    abm = RunMacro("Get ABM Manager", Args)

    // Process tours
    arrsOut = null
    purposes = {'Work1', 'Work2', 'Univ1', 'Univ2', 'School'}
    for p in purposes do
        // Process direct tours
        spec = {Label: p, abmManager: abm, Periods: Args.TimePeriods}
        arrs = RunMacro("Synthesize Tour Data", spec)
        RunMacro("Append Arrays", arrs, &arrsOut)
    end

    // Write Tours Table (from the direct tours above)
    vwTours = RunMacro("Write Tours Table", {Data: arrsOut})

    // Modify tours that also had PUDO activities. This will update the appropriate records of the tour file.
    periods = {"AM", "PM", "OP"}
    skimArgs = {TimePeriods: Args.TimePeriods}
    for p in periods do
        mObj.(p) = CreateObject("Matrix", Args.("HighwaySkim" + p))
        mObj.(p).SetIndex({RowIndex: "TAZ", ColIndex: "TAZ"})
        skimArgs.(p) = mObj.(p).Time
    end
    RunMacro("Modify Tours for PUDO", {ToursView: vwTours, abmManager: abm, SkimArgs: skimArgs, Periods: Args.TimePeriods})
    
    // Add Separate PUDO Tours. This will create two new tables that have to be concatenated with the tours file.
    puFile = GetTempPath() + "PUTours.bin"
    doFile = GetTempPath() + "DOTours.bin"
    opts = {abmManager: abm, SkimArgs: skimArgs, Periods: Args.TimePeriods, DropoffTourFile: doFile, PickupTourFile: puFile}
    RunMacro("Add Separate PUDO tours", opts)
    
    // Export tours table to temporary file
    tempToursFile = GetTempPath() + "TempTours.bin"
    ExportView(vwTours + "|", "FFB", tempToursFile,,)
    CloseView(vwTours)
    
    // Concatenate PU, DO and ToursFile
    ConcatenateFiles({tempToursFile, doFile, puFile}, Args.MandatoryTours)
    dictFile = Substitute(Args.MandatoryTours, ".bin", ".dcb",)
    CopyFile(GetTempPath() + "TempTours.dcb", dictFile)

    // Postprocess: Fill Tour ID field
    vwT = OpenTable("Tours", "FFB", {Args.MandatoryTours})
    v = GetDataVector(vwT + "|", "TourID",)
    v1 = Vector(v.length, "Long", {{"Sequence", 1, 1}})
    
    sortOrder = {{'HID', 'Ascending'}, {'PerID', 'Ascending'}, {'ActivityStartTime', 'Ascending'}}
    SetDataVector(vwT + "|", "TourID", v1, {SortOrder: sortOrder})
    CloseView(vwT)
    Return(true)
endMacro


/*
    Generate tours from person decision for a given purpose: 'Work1', 'Work2', 'HigherEd1', 'HigherEd2', 'School'
    Returns an option array. The option names are the field names and the values are vectors containing one record for each tour.
*/
Macro "Synthesize Tour Data"(spec)
    label = spec.Label
    pInfo = RunMacro("Get Purpose Info", label)
    purpose = pInfo.Purpose
    tourNumber = pInfo.TourNo
    abm = spec.abmManager
    PeriodInfo = spec.Periods

    destFld = purpose + "TAZ" // e.g. 'WorkTAZ'
    actStartFld = label + "_StartTime"
    actDurFld = label + "_Duration"
    ttForwardFld = "HomeTo" + purpose + "Time"
    ttReturnFld = purpose + "ToHomeTime"
    if purpose = "School" then do
        modeFldF = "SchoolForwardMode"
        modeFldR = "SchoolReturnMode"
        purpFilter = "(AttendSchool = 1 or AttendDaycare = 1)"
    end
    else do
        modeFldF = purpose + "Mode"
        modeFldR = purpose + "Mode"
        purpFilter = printf("(Number%sTours >= %s)", {purpose, tourNumber})
    end
    modeCodeFldF = modeFldF + "Code"
    modeCodeFldR = modeFldR + "Code"
    
    // Get Filter: Example 'NumberWorkTours = 1 and WorkTAZ <> null'
    filter = printf("%s and %s <> null and %s_StartTime <> null", {purpFilter, destFld, label})
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})
    nRecs = set.Size
    
    hhIDSpec = GetFieldFullSpec(abm.HHView, "HouseholdID")
    flds = {"PersonID", hhIDSpec, "TAZID", destFld, actStartFld, actDurFld, ttForwardFld,
            modeFldF, modeFldR, modeCodeFldF, modeCodeFldR, 
            "DropoffTourFlag", "PickupTourFlag", "NDropOffsEnRoute", "NPickupsEnRoute"}
    vecs = abm.GetPersonHHVectors(flds)

    // Tours
    arrs = null
    arrs.PerID = v2a(vecs.PersonID)
    arrs.HID = v2a(vecs.(hhIDSpec))
    arrs.TourType = v2a(Vector(nRecs, "String", {Constant: "Mandatory"}))
    arrs.TourPurpose = v2a(Vector(nRecs, "String", {Constant: purpose}))
    arrs.MandTourNo = arrTourNo
    arrs.HTAZ = v2a(vecs.TAZID)
    arrs.Origin = v2a(vecs.TAZID)
    arrs.Destination = v2a(vecs.(destFld))
    arrs.ForwardMode = v2a(vecs.(modeFldF))
    arrs.ReturnMode = v2a(vecs.(modeFldR))
    arrs.ForwardModeCode = v2a(vecs.(modeCodeFldF))
    arrs.ReturnModeCode = v2a(vecs.(modeCodeFldR))
    
    vTourStart = vecs.(actStartFld) - vecs.(ttForwardFld)
    arrs.TourStartTime = v2a(vTourStart)
    arrs.DestArrTime = v2a(vecs.(actStartFld))
    arrs.ActivityStartTime = v2a(vecs.(actStartFld))
    arrs.ActivityDuration = v2a(vecs.(actDurFld))
    
    vActEnd = vecs.(actStartFld) + vecs.(actDurFld)
    arrs.ActivityEndTime = v2a(vActEnd)
    arrs.DestDepTime = v2a(vActEnd)
    
    //vTourEnd = vActEnd + vecs.(ttReturnFld)
    vTourEnd = vActEnd + vecs.(ttForwardFld)
    arrs.TourEndTime = v2a(vTourEnd)
    
    vTODF = RunMacro("Get TOD Vector", vTourStart, PeriodInfo)
    vTODR = RunMacro("Get TOD Vector", vActEnd, PeriodInfo)
    arrs.TODForward = v2a(vTODF)
    arrs.TODReturn = v2a(vTODR)

    // Identify tour records that need to be modified for PUDO
    vModifyDO = if (vecs.NDropOffsEnRoute > 0) or (vecs.DropoffTourFlag = 'W1' and vecs.(modeCodeFldF) = 2) then 1
    arrs.ModifyRecforDropoff = v2a(vModifyDO)

    vModifyPU = if (vecs.NPickupsEnRoute > 0) or (vecs.PickupTourFlag = 'W1' and vecs.(modeCodeFldR) = 2) then 1
    arrs.ModifyRecforPickup = v2a(vModifyPU)

    Return(arrs)
endMacro


Macro "Get Purpose Info"(str)
    validStrs = {'Work1', 'Work2', 'Univ1', 'Univ2', 'School'}
    if validStrs.position(str) = 0 then
        Throw("Invalid input to macro 'Get Purpose Info'")

    if lower(str) = "school" or lower(str) = "daycare" then do
        purpose = "School"
        tourNo = "1"
    end
    else do
        purpose = Left(str, StringLength(str) - 1)
        tourNo = Right(str, 1)
    end
    Return({Purpose: purpose, TourNo: tourNo})
endMacro


/*
    Write data from option array into the final tours table
*/
Macro "Write Tours Table"(spec)
    data = spec.Data
    nRecs = data.Origin.length
    vwOut = RunMacro("Create Empty Tour File", {ViewName: "Tours", NRecords: nRecs})
    
    vecsOut = null
    for item in data do
        fld = item[1]
        vecsOut.(fld) = a2v(data.(fld))
    end
    vecsOut.AssignForwardHalf = if vecsOut.ForwardMode = 'Carpool' and vecsOut.TourPurpose = 'School' then 0 else 1
    vecsOut.AssignReturnHalf = if vecsOut.ReturnMode = 'Carpool' and vecsOut.TourPurpose = 'School' then 0 else 1
    SetDataVectors(vwOut + "|", vecsOut,)

    // Fill Mandatory Tour No field
    obj = CreateObject("Table", vwOut)
    order = {{"PerID", "Ascending"}, {"ActivityStartTime", "Ascending"}}
    obj.Sort({FieldArray: order})
    
    vP = obj.PerID
    vPrevP = RunMacro("Shift Vector", {Vector: vP, Method: 'Prev'})
    v = if (vP = vPrevP) then 2 else 1
    obj.MandTourNo = v

    Return(vwOut)
endMacro


// Create a temporary in-memory tour table and return the view
Macro "Create Empty Tour File"(spec)
    flds = {{"TourID", "Integer", 12, null, "Yes"},
            {"HID", "Integer", 12, null, "Yes"},
            {"PerID", "Integer", 12, null, "Yes"},
            {"TourType", "String", 12, null, "No"},
            {"TourPurpose", "String", 8, null, "Yes"},
            {"MandTourNo", "Short", 2, null, "No"},
            {"HTAZ", "Integer", 12, null, "Yes"},
            {"Origin", "Integer", 12, null, "Yes"},
            {"Destination", "Integer", 12, null, "Yes"},
            {"ForwardModeCode", "Integer", 2, null, "No"},
            {"ReturnModeCode", "Integer", 2, null, "No"},
            {"ForwardMode", "String", 15, null, "No"},
            {"ReturnMode", "String", 15, null, "No"},
            {"TourStartTime", "Integer", 12, null, "No"},
            {"DestArrTime", "Integer", 12, null, "No"},
            {"ActivityStartTime", "Integer", 12, null, "No"},
            {"ActivityDuration", "Integer", 12, null, "No"},
            {"ActivityEndTime", "Integer", 12, null, "No"},
            {"DestDepTime", "Integer", 12, null, "No"},
            {"TourEndTime", "Integer", 12, null, "No"},
            {"TODForward", "String", 2, null, "No"},
            {"TODReturn", "String", 2, null, "No"},
            {"NumDropoffs", "Short", 2, null, "No"},
            {"DropoffKidID1", "Integer", 12, null, "No"},
            {"DropoffTAZ1", "Integer", 12, null, "No"},
            {"ArrDropoff1", "Integer", 12, null, "No"},
            {"DepDropoff1", "Integer", 12, null, "No"},
            {"DropoffKidID2", "Integer", 12, null, "No"},
            {"DropoffTAZ2", "Integer", 12, null, "No"},
            {"ArrDropoff2", "Integer", 12, null, "No"},
            {"DepDropoff2", "Integer", 12, null, "No"},
            {"NumPickups", "Short", 2, null, "No"},
            {"PickupKidID1", "Integer", 12, null, "No"},
            {"PickupTAZ1", "Integer", 12, null, "No"},
            {"ArrPickup1", "Integer", 12, null, "No"},
            {"DepPickup1", "Integer", 12, null, "No"},
            {"PickupKidID2", "Integer", 12, null, "No"},
            {"PickupTAZ2", "Integer", 12, null, "No"},
            {"ArrPickup2", "Integer", 12, null, "No"},
            {"DepPickup2", "Integer", 12, null, "No"},
            {"StopsChoice", "String", 3, null, "No"},
            {"NForwardStops", "Short", 2, null, "No"},
            {"NReturnStops", "Short", 2, null, "No"},
            {"StopForwardTAZ", "Integer", 12, null, "No"},
            {"StopReturnTAZ", "Integer", 12, null, "No"},
            {"IsStopBeforeDropoffs", "Short", 2, null, "No"},
            {"IsStopBeforePickups", "Short", 2, null, "No"},
            {"ForwardStopDeltaTT", "Real", 12, 2, "No"},
            {"ReturnStopDeltaTT", "Real", 12, 2, "No"},
            {"ForwardStopDurChoice", "String", 12,, "No"},
            {"ForwardStopDuration", "Real", 12, 2, "No"},
            {"ReturnStopDurChoice", "String", 12,, "No"},
            {"ReturnStopDuration", "Real", 12, 2, "No"},
            {"NumSubTours", "Short", 2, null, "No"},
            {"ModifyRecforDropoff", "Short", 2, null, "No"},
            {"ModifyRecforPickup", "Short", 2, null, "No"},
            {"AssignForwardHalf", "Short", 2, null, "No"},
            {"AssignReturnHalf", "Short", 2, null, "No"}}
    vwOut = CreateTable(spec.ViewName,, "MEM", flds)
    if spec.NRecords > 0 then
        AddRecords(vwOut,,, {{"Empty Records", spec.NRecords}})    
    Return(vwOut)
endMacro


Macro "Append Arrays"(arrs, arrsOut)
    if arrsOut = null then
        arrsOut = CopyArray(arrs)
    else do
        inputFlds = arrs.Map(do (f) Return(f[1]) end) // Get input option names
        inputLength = arrs[1][2].length
        
        outputFlds = arrsOut.Map(do (f) Return(f[1]) end) // Get output option names
        outputLength = arrsOut[1][2].length
        
        // Loop over each output field and append appropriate vector
        for fld in outputFlds do
            if inputFlds.position(fld) > 0 then // Append given input array
                arrsOut.(fld) = arrsOut.(fld) + CopyArray(arrs.(fld))
            else do                             // Not present in input arrays. Append null records.
                dim nullArray[inputLength]
                arrsOut.(fld) = CopyArray(arrsOut.(fld)) + nullArray
            end
        end

        // Check for any new input fields not already present in arrsOut
        for fld in inputFlds do
            if outputFlds.position(fld) = 0 then do // New field
                dim nullArray[outputLength]
                arrsOut.(fld) = nullArray + CopyArray(arrs.(fld))
            end
        end        
    end
endMacro
