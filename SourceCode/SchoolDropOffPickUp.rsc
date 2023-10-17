/*
    Allocate potential pick up and drop off person for the school kids
    Only kids where both pickup/drop off persons allocated can choose carpool
    Allocate to non workers with cars and license or to mandatory tour persons (if school and mandatory tour schedules do not clash)
*/
Macro "School Carpool Eligibility"(abm)
    abm.ActivatePersonSet()
    abm.[Person.NDropoffs] = 0 // Clear field
    abm.[Person.NPickups] = 0  // Clear field

    // Fill formula field for adults eligible for PUDO
    ageFilter = "(Age > 18 and Age <= 70)"                                      // Only select persons in this age range 
    noUnivFilter = "(AttendUniv <> 1)"                                          // Only select persons who are not univ students or school students
    noTwoTours = "(nz(NumberWorkTours) < 2)"                                    // Only select persons who make 0 or 1 work tours
    vehAvailFilter = "((VehicleUsed = 1 or VehiclesRem > 0) and License = 1)"   // Select persons who have a license with available car
    modeFilter = "(!(Lower(WorkMode) contains 'bus') and !(Lower(WorkMode) contains 'rail') and Lower(WorkMode) <> 'carpool')"             // No PT or Carpool

    exprStr = printf("if %s and %s and %s and %s and %s then 1 else 0", {ageFilter, noUnivFilter, noTwoTours, vehAvailFilter, modeFilter})
    expr = CreateExpression(abm.PersonHHView, "EligibleAdult", exprStr,)
    
    // Select persons eligible for drop-off and pick-up. Also select kids who attend school/daycare.
    hhFilter = "(NSchoolKids > 0)"                               // Only select persons in HHs that have school kids attending school
    kidFilter = "(AttendSchool = 1 or AttendDaycare = 1)"        // Make sure to choose school/daycare students in this list
    personFilter = "(EligibleAdult = 1)"                         // Select eligible adults
    filter = printf("%s and (%s or %s)", {hhFilter, kidFilter, personFilter})
    set = abm.CreatePersonSet({Filter: filter, Activate: 1, UsePersonHHView: 1})

    // Allocate drop offs
    // For each student who attends school/daycare, identify if drop off is possible and if so associate the person who will drop off
    // Identify if the drop off will be part of a separate tour ('S'), will on the way to work on the first work tour ('W1') or
    // will be on the way to work for the second work tour ('W2')
    // Only kids who have a drop off person associated will be allowed to choose 'Carpool' as a school forward mode
    RunMacro("Allocate Dropoffs", abm)
    
    // Allocate pick ups
    // For each student who attends school/daycare, identify if pick up is possible and if so associate the person who will pick up
    // Identify if the pick up will be part of a separate tour ('S'), will on the way back from first work tour ('W1') or
    // will be on the way back from the second work tour ('W2')
    // Only kids who have a pick up person associated will be allowed to choose 'Carpool' as a school return mode
    RunMacro("Allocate Pickups", abm)

    DestroyExpression(GetFieldFullSpec(abm.PersonHHView, expr))
endMacro

/* 
    Allocate dropoff person for school carpool students
    Consider only HHs where kids make school trips
    Priority of assignment:
    1. Non workers and non univ students with (Age > 18 and Age <= 70) with license and available car 
    2. Worker/Univ student with license and car but no mandatory tour on the day
    3. Worker/Univ student with license and car but whose work/univ schedule does not overlap with school schedule (for e.g. school drop off in the AM peak but work start time is PM)
    4. Worker/Univ student who DriveAlone to work/univ with small adjustments to their schedule to facilitate pickup/dropoff

    Dropoff allocation independent of pickup allocation. Therefore person dropping off and picking up could be different.
    Do not allow more than 2 dropoffs/pickups for each person.
*/
Macro "Allocate Dropoffs"(abm)
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "Allocate Dropoff", 
                InputFields: {"VehicleUsed", "NSchoolKids", "WorkMode", "NumberWorkTours", "EligibleAdult",
                              "Work1_StartTime", "School_StartTime", "Work1_Duration", "AttendSchool", "AttendDaycare",
                              "HometoWorkTime", "WorktoHomeTime", "HometoSchoolTime", "PersonID"},
                OutputFields: {"DropoffPersonID", "DropoffTourFlag", "NDropOffs"},
                SortOrder: {AttendDaycare: "Descending", AttendSchool: "Descending", NumberWorkTours: "Descending", 
                            WorkerCategory: "Ascending", VehicleUsed: "Ascending", Age: "Ascending"}
                }
    abm.Iterate(iterOpts)
endMacro


Macro "Allocate Dropoff"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    nCarpool = vecs.NSchoolKids[startIdx] // Get number of potential kids who need carpool from the first record in the set
    maxDropoffs = 2
    lateAllowance = 15      // Allow person to be late to work at most by this amount
    earlyAllowance = 30     // Allow person to be early to work at most by this amount
    buffer = 5              // Allow at least so much time between tours

    for i = 1 to nCarpool do // For each carpool kid
        kidIdx = startIdx + (i - 1)
        schStart = vecs.School_StartTime[kidIdx]
        h2sTT = vecs.HometoSchoolTime[kidIdx]
        schDep = schStart - h2sTT
        doAssigned = 0
        for j = (startIdx + nCarpool) to endIdx do // Loop over each eligible adult. j is index ID of adult being processed.
            if vecs.EligibleAdult[j] = 0 then                
                continue
            
            if vecsOut.NDropoffs[j] >= maxDropoffs then        // Already equals max dropoffs. Process next adult.
                continue

            if nz(vecs.NumberWorkTours[j]) = 0 then do // Non worker or someone who makes no mandatory tours. Assign and move on.
                doAssigned = 1
                tourFlag = 'S' // separate drop off tour
            end
            else do // Worker
                work1Dep = vecs.Work1_StartTime[j] - vecs.HometoWorkTime[j]
                work1Ret = vecs.Work1_StartTime[j] + vecs.Work1_Duration[j] + vecs.WorktoHomeTime[j]
                wrkMode = vecs.WorkMode[j]
                if schDep + 2*h2sTT + buffer <= work1Dep then do             // School departure sufficiently earlier than work tour 1 departure
                    doAssigned = 1
                    tourFlag = 'S' // separate drop off tour
                end
                else if wrkMode = 'DriveAlone' and (schDep >= (work1Dep - earlyAllowance) and schDep <= (work1Dep + lateAllowance)) then do // School departure closer to Work Tour 1 departure. Drop off stop during first work tour.
                    doAssigned = 1
                    tourFlag = 'W1' // Drop off on first work tour
                end
                else if schDep >= work1Ret + buffer then do // School dep after work 1 return. Possible unless second work tour interferes
                    doAssigned = 1
                    tourFlag = 'S' // separate drop off tour
                end
            end

            if doAssigned = 1 then  do  // Finalise dropoff for student. Increment drop off counter for adult.
                vecsOut.DropoffPersonID[kidIdx] = vecs.PersonID[j]
                vecsOut.DropoffTourFlag[kidIdx] = tourFlag
                vecsOut.NDropoffs[j] = vecsOut.NDropoffs[j] + 1
            end

            if doAssigned = 1 then // Process next kid
                break
        end // Adult Loop
    end // Kid Loop
endMacro


Macro "Allocate Pickups"(abm)
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "Allocate Pickup", 
                InputFields: {"VehicleUsed", "NSchoolKids", "WorkMode", "NumberWorkTours", "EligibleAdult",
                              "Work1_StartTime", "School_StartTime",
                              "Work1_Duration", "School_Duration", "AttendSchool", "AttendDaycare",
                              "HometoWorkTime", "WorktoHomeTime", "HometoSchoolTime", "PersonID"},
                OutputFields: {"PickupPersonID", "PickupTourFlag", "NPickups"},
                SortOrder: {AttendDaycare: "Descending", AttendSchool: "Descending", NumberWorkTours: "Descending", 
                            WorkerCategory: "Ascending", VehicleUsed: "Ascending", Age: "Ascending"}
                }
    abm.Iterate(iterOpts)
endMacro


Macro "Allocate Pickup"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    nCarpool = vecs.NSchoolKids[startIdx] // Get number of kids who need carpool from the first record in the set
    maxPickups = 2
    lateAllowance = 15      // Allow person to be late to work (or leave work early) at most by this amount
    earlyAllowance = 30     // Allow person to be early to work at most by this amount
    buffer = 5              // Allow at least so much time between tours

    for i = 1 to nCarpool do // For each carpool kid
        kidIdx = startIdx + (i - 1)
        schEnd = vecs.School_StartTime[kidIdx] + vecs.School_Duration[kidIdx]
        h2sTT = vecs.HometoSchoolTime[kidIdx]
        s2hTT = h2sTT
        schRet = schEnd + s2hTT
        puAssigned = 0
        for j = (startIdx + nCarpool) to endIdx do // Loop over each eligible adult. j is index ID of adult being processed.
            if vecs.EligibleAdult[j] = 0 then                
                continue

            if vecsOut.NPickups[j] >= maxPickups then        // Equals max pickups. Skip to next adult.
                continue

            if nz(vecs.NumberWorkTours[j]) = 0 then do // Non worker or someone who makes no mandatory tours. Assign and move on.
                puAssigned = 1
                tourFlag = 'S'
            end
            else do // Worker
                h2wTT = vecs.HometoWorkTime[j]
                w2hTT = vecs.WorktoHomeTime[j]
                work1Dep = vecs.Work1_StartTime[j] - h2wTT
                work1Ret = vecs.Work1_StartTime[j] + vecs.Work1_Duration[j] + w2hTT
                work1EarlyRet = work1Ret - lateAllowance        // Arrive early from work by the lateAllowance time
                work1LateRet = work1Ret + earlyAllowance        // Max delayed return from work to accomodate pickup
                wrkMode = vecs.WorkMode[j]
                
                // Note that the time between work1Dep and work1EarlyRet is assumed to be the non negotiable time.
                // If the school ends within this interval, then person cannot pickup
                if schRet + buffer <= work1Dep then do  // School ends before or sufficiently close to departure to first work tour, make separate PU tour
                    puAssigned = 1
                    tourFlag = 'S'
                end
                else if schEnd >= work1Ret + h2sTT then do
                    puAssigned = 1
                    tourFlag = 'S'
                end
                else if schEnd < work1EarlyRet + h2sTT then // Not possible since school ends during the non negotiable window where person is at work
                    puAssigned = 0
                else if (schEnd < work1LateRet + h2sTT) and (wrkMode = 'DriveAlone') then do // School end times enable pick up directly from work
                    puAssigned = 1
                    tourFlag = 'W1'
                end
            end

            if puAssigned = 1 then  do // Finalise pickup for student. Increment pickup counter for adult.
                vecsOut.PickupPersonID[kidIdx] = vecs.PersonID[j]
                vecsOut.PickupTourFlag[kidIdx] = tourFlag
                vecsOut.NPickups[j] = vecsOut.NPickups[j] + 1
            end

            if puAssigned = 1 then // Process next kid
                break
        end
    end
endMacro


/* 
    This macro modifies the mandatory tour records for workers who make PUDO stops on their main tour
    It also updates the mandatory tour records of the kids being dropped off
*/
Macro "Modify Tours for PUDO"(spec)
    RunMacro("Modify Tours for Dropoff", spec)
    RunMacro("Modify Tours for Pickup", spec)
endMacro


Macro "Modify Tours for Dropoff"(spec)
    abm = spec.abmManager
    vwTours = spec.ToursView

    // Select kids who chose carpool and who are are dropped/picked on enroute by adults. Also select adults making these stops.
    // Note that persons who make more than 1 mandatory tour during the day were automatically disqualified from PUDO
    filter = "(DropoffTourFlag = 'W1' and SchoolForwardModeCode = 2) or (NDropoffsEnroute > 0)"
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})
    if set.Size = 0 then
        Return()

    // Loop over records (by personID and write one record for each person in a temporary table)
    flds = {{"ID", "Integer", 12, null, "Yes"},
            {"HomeDep", "Integer", 12, null, "No"},
            {"KidID1", "Integer", 12, null, "No"},
            {"TAZ1", "Integer", 12, null, "No"},
            {"ArrTAZ1", "Integer", 12, null, "No"},
            {"DepTAZ1", "Integer", 12, null, "No"},
            {"KidID2", "Integer", 12, null, "No"},
            {"TAZ2", "Integer", 12, null, "No"},
            {"ArrTAZ2", "Integer", 12, null, "No"},
            {"DepTAZ2", "Integer", 12, null, "No"},
            {"NPUDO", "Integer", 12, null, "No"},
            {"SchoolArr", "Integer", 12, null, "No"},
            {"WorkArr", "Integer", 12, null, "No"}}
    vwTemp = CreateTable("TempPUDO",, "MEM", flds)
    AddRecords(vwTemp,,, {"Empty Records": set.Size})
    {flds, specs} = GetFields(vwTemp,)
    
    exprStr = printf("if AttendSchool = 1 or AttendDaycare = 1 then DropoffPersonID else %s", {abm.PersonID})
    adultIDFld = CreateExpression(abm.PersonHHView, "DesignatedAdultID", exprStr,)
    
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "ModifyDropoff",
                MacroArgs: {SkimArgs: spec.SkimArgs},
                LoopOn: adultIDFld,
                InputFields: {adultIDFld, "SchoolTAZ", "TAZID", "School_StartTime", "PersonID", "WorkTAZ"},
                OutputView: vwTemp,
                SortOrder: {{adultIDFld, 'Ascending'}, {"School_StartTime", 'Ascending'}, {'Age', 'Ascending'}}
                }
    abm.Iterate(iterOpts)
    DestroyExpression(GetFieldFullSpec(abm.PersonHHView, adultIDFld))

    // Join temporary file to the main tour file and update values
    vwJ = JoinViews("Temp_Tours", GetFieldFullSpec(vwTemp, "ID"), GetFieldFullSpec(vwTours, "PerID"),)
    vecs = GetDataVectors(vwJ + "|", flds + {"ModifyRecforDropoff", "TourStartTime", "DestArrTime"}, {OptArray: 1})
    vecsSet = null
    vecsSet.DropoffKidID1 = vecs.KidID1
    vecsSet.DropoffKidID2 = vecs.KidID2
    vecsSet.DropoffTAZ1 = vecs.TAZ1
    vecsSet.DropoffTAZ2 = vecs.TAZ2
    vecsSet.ArrDropoff1 = vecs.ArrTAZ1
    vecsSet.ArrDropoff2 = vecs.ArrTAZ2
    vecsSet.DepDropoff1 = vecs.DepTAZ1
    vecsSet.DepDropoff2 = vecs.DepTAZ2
    vecsSet.NumDropoffs = vecs.NPUDO
    vecsSet.TourStartTime = if vecs.ModifyRecforDropoff <> 1 then vecs.TourStartTime
                                else vecs.HomeDep
    vecsSet.DestArrTime = if vecs.ModifyRecforDropoff <> 1 then vecs.DestArrTime
                            else if vecs.NPUDO > 0 then vecs.WorkArr 
                            else vecs.SchoolArr
    SetDataVectors(vwJ + "|", vecsSet,)
    CloseView(vwJ)
    CloseView(vwTemp)
endMacro



Macro "Modify Tours for Pickup"(spec)
    abm = spec.abmManager
    vwTours = spec.ToursView
    
    // Select kids who chose carpool and who are are dropped/picked on enroute by adults. Also select adults making these stops.
    // Note that persons who make more than 1 mandatory tour during the day were automatically disqualified from PUDO
    filter = "(PickupTourFlag = 'W1' and SchoolReturnModeCode = 2) or (NPickupsEnroute > 0)"
    set = abm.CreatePersonSet({Filter: filter, Activate: 1})
    if set.Size = 0 then
        Return()

    // Loop over records (by personID and write one record for each person in a temporary table)
    flds = {{"ID", "Integer", 12, null, "Yes"},
            {"KidID1", "Integer", 12, null, "No"},
            {"TAZ1", "Integer", 12, null, "No"},
            {"ArrTAZ1", "Integer", 12, null, "No"},
            {"DepTAZ1", "Integer", 12, null, "No"},
            {"KidID2", "Integer", 12, null, "No"},
            {"TAZ2", "Integer", 12, null, "No"},
            {"ArrTAZ2", "Integer", 12, null, "No"},
            {"DepTAZ2", "Integer", 12, null, "No"},
            {"NPUDO", "Integer", 12, null, "No"},
            {"SchoolArr", "Integer", 12, null, "No"},
            {"SchoolDep", "Integer", 12, null, "No"},
            {"WorkDep", "Integer", 12, null, "No"},
            {"HomeArr", "Integer", 12, null, "No"}}
    vwTemp = CreateTable("TempPUDO",, "MEM", flds)
    AddRecords(vwTemp,,, {"Empty Records": set.Size})
    {flds, specs} = GetFields(vwTemp,)
    
    schEndFld = CreateExpression(abm.PersonHHView, "School_EndTime", "School_StartTime + School_Duration",)
    exprStr = printf("if AttendSchool = 1 or AttendDaycare = 1 then PickupPersonID else %s", {abm.PersonID})
    adultIDFld = CreateExpression(abm.PersonHHView, "DesignatedAdultID", exprStr,)
    
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "ModifyPickup",
                MacroArgs: {SkimArgs: spec.SkimArgs},
                LoopOn: adultIDFld,
                InputFields: {adultIDFld, "SchoolTAZ", "TAZID", schEndFld, "PersonID", "WorkTAZ"},
                OutputView: vwTemp,
                SortOrder: {{adultIDFld, 'Ascending'}, {schEndFld, 'Ascending'}, {'Age', 'Ascending'}}
                }
    abm.Iterate(iterOpts)
    DestroyExpression(GetFieldFullSpec(abm.PersonHHView, adultIDFld))

    // Join temporary file to the main tour file and update values
    vwJ = JoinViews("Temp_Tours", GetFieldFullSpec(vwTemp, "ID"), GetFieldFullSpec(vwTours, "PerID"),)
    vecs = GetDataVectors(vwJ + "|", flds + {"ModifyRecForPickup", "DestDepTime", "TourEndTime"}, {OptArray: 1})
    vecsSet = null
    vecsSet.PickupKidID1 = vecs.KidID1
    vecsSet.PickupKidID2 = vecs.KidID2
    vecsSet.PickupTAZ1 = vecs.TAZ1
    vecsSet.PickupTAZ2 = vecs.TAZ2
    vecsSet.ArrPickup1 = vecs.ArrTAZ1
    vecsSet.ArrPickup2 = vecs.ArrTAZ2
    vecsSet.DepPickup1 = vecs.DepTAZ1
    vecsSet.DepPickup2 = vecs.DepTAZ2
    vecsSet.NumPickups = vecs.NPUDO
    vecsSet.DestDepTime = if vecs.ModifyRecForPickup <> 1 then vecs.DestDepTime
                            else if vecs.NPUDO > 0 then vecs.WorkDep 
                            else vecs.SchoolDep
    vecsSet.TourEndTime = if vecs.ModifyRecForPickup <> 1 then vecs.TourEndTime 
                            else vecs.HomeArr
    SetDataVectors(vwJ + "|", vecsSet,)
    CloseView(vwJ)
    CloseView(vwTemp)
    DestroyExpression(GetFieldFullSpec(abm.PersonHHView, schEndFld))
endMacro


Macro "ModifyDropoff"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    i = spec.RecordIndex
    SkimArgs = spec.MacroArgs.SkimArgs
    doffTime = 3 // Time to drop off kids
    
    adultIdx = startIdx
    adultID = vecs.DesignatedAdultID[startIdx]
    if endIdx - startIdx = 1 then do // Simple case. Only one kid dropped off.
        // Get TT from home to school zone
        kidIdx = endIdx
        schStart = vecs.School_StartTime[kidIdx]
        h2stt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[kidIdx], Dest: vecs.SchoolTAZ[kidIdx], DepTime: schStart})
        startTime = schStart - h2stt // Can safely assume this does not cross into previous day
        
        // Kid record
        vecsOut.ID[i] = vecs.PersonID[kidIdx]
        vecsOut.HomeDep[i] = startTime
        vecsOut.SchoolArr[i] = schStart
        
        // Adult record
        i = i + 1
        vecsOut.ID[i] = adultID
        vecsOut.KidID1[i] = vecs.PersonID[kidIdx]
        vecsOut.HomeDep[i] = startTime
        vecsOut.ArrTAZ1[i] = schStart
        vecsOut.DepTAZ1[i] = schStart + doffTime // Allow extra min for drop-off
        vecsOut.TAZ1[i] = vecs.SchoolTAZ[kidIdx]
        s2w_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.SchoolTAZ[kidIdx], Dest: vecs.WorkTAZ[adultIdx], DepTime: schStart + doffTime})
        vecsOut.WorkArr[i] = schStart + doffTime + s2w_tt
        vecsOut.NPUDO[i] = 1
        
        nRecsAdded = 2
    end
    else do // Two kids. Note order is by dult and then by kids in ascending school start time, followed by age
        kidIdx1 = startIdx + 1
        kidIdx2 = endIdx
        schStart1 = vecs.School_StartTime[kidIdx1]
        schTAZ1 = vecs.SchoolTAZ[kidIdx1]
        schStart2 = vecs.School_StartTime[kidIdx2]
        schTAZ2 = vecs.SchoolTAZ[kidIdx2]
        hs1_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[kidIdx1], Dest: schTAZ1, DepTime: schStart1})
        s1s2_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ1, Dest: schTAZ2, DepTime: schStart1})
        if (schStart2 - schStart1) < s1s2_tt + doffTime then do // Need to make up time by leaving home even earlier. First kid will be early to school.
            deltaTT = s1s2_tt - (schStart2 - schStart1) + doffTime
            tourStart = schStart1 - hs1_tt - deltaTT
        end
        else
            tourStart = schStart1 - hs1_tt
        
        // Kid 1 record (1st dropoff)
        vecsOut.ID[i] = vecs.PersonID[kidIdx1]
        vecsOut.HomeDep[i] = tourStart
        arrTAZ1 = tourStart + hs1_tt
        vecsOut.SchoolArr[i] = arrTAZ1

        // Kid 2 record (2nd dropoff)
        i = i + 1
        vecsOut.ID[i] = vecs.PersonID[kidIdx2]
        vecsOut.HomeDep[i] = tourStart
        //vecsOut.ArrTAZ1[i] = arrTAZ1
        //vecsOut.DepTAZ1[i] = arrTAZ1 + doffTime
        //vecsOut.TAZ1[i] = vecs.SchoolTAZ[kidIdx1]
        arrTAZ2 = arrTAZ1 + doffTime + s1s2_tt
        vecsOut.SchoolArr[i] = arrTAZ2

        // Adult record
        i = i + 1
        vecsOut.ID[i] = adultID
        vecsOut.KidID1[i] = vecs.PersonID[kidIdx1]
        vecsOut.KidID2[i] = vecs.PersonID[kidIdx2]
        vecsOut.HomeDep[i] = tourStart
        vecsOut.ArrTAZ1[i] = arrTAZ1
        vecsOut.DepTAZ1[i] = arrTAZ1 + doffTime
        vecsOut.TAZ1[i] = vecs.SchoolTAZ[kidIdx1]
        vecsOut.ArrTAZ2[i] = arrTAZ2
        vecsOut.DepTAZ2[i] = arrTAZ2 + doffTime
        vecsOut.TAZ2[i] = vecs.SchoolTAZ[kidIdx2]
        s2w_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ2, Dest: vecs.WorkTAZ[adultIdx], DepTime: arrTAZ2 + doffTime})
        vecsOut.WorkArr[i] = arrTAZ2 + doffTime + s2w_tt
        vecsOut.NPUDO[i] = 2

        nRecsAdded = 3
    end
    Return(nRecsAdded)
endMacro


Macro "ModifyPickup"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    i = spec.RecordIndex
    SkimArgs = spec.MacroArgs.SkimArgs
    
    // Process
    puTime = 3
    adultIdx = startIdx
    adultID = vecs.DesignatedAdultID[startIdx]
    kidIdx1 = startIdx + 1
    schEnd1 = vecs.School_EndTime[kidIdx1]
    schTAZ1 = vecs.SchoolTAZ[kidIdx1]
    ws1ttEst = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.WorkTAZ[adultIdx], Dest: schTAZ1, DepTime: schEnd})
    ws1_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.WorkTAZ[adultIdx], Dest: schTAZ1, DepTime: schEnd - ws1ttEst}) // Improve estimate
    workDep = schEnd1 - ws1_tt
    arrTAZ1 = schEnd1
    depTAZ1 = schEnd1 + puTime
    
    if (endIdx - startIdx) = 1 then do // Simple case. Only one kid picked up.
        // Kid record
        s1h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ1, Dest: vecs.TAZID[kidIdx1], DepTime: depTAZ1})
        homeArr = depTAZ1 + s1h_tt
        vecsOut.ID[i] = vecs.PersonID[kidIdx1]
        vecsOut.SchoolDep[i] = depTAZ1
        vecsOut.HomeArr[i] = homeArr
        
        // Adult record
        i = i + 1
        vecsOut.ID[i] = adultID
        vecsOut.KidID1[i] = vecs.PersonID[kidIdx1]
        vecsOut.WorkDep[i] = workDep
        vecsOut.ArrTAZ1[i] = arrTAZ1
        vecsOut.DepTAZ1[i] = depTAZ1
        vecsOut.TAZ1[i] = schTAZ1
        vecsOut.HomeArr[i] = homeArr
        vecsOut.NPUDO[i] = 1
        
        nRecsAdded = 2
    end
    else do // Two kids. Note order is by ascending school end time, followed by age
        kidIdx2 = endIdx
        schEnd2 = vecs.School_EndTime[endIdx]
        schTAZ2 = vecs.SchoolTAZ[endIdx]
        s1s2_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ1, Dest: schTAZ2, DepTime: depTAZ1})
        
        // For pickup, there is no concept of making up time because kids cannot leave before school ends. 
        // Therefore adult can be late for 2nd pickup.
        
        // Kid 1 record (1st pickup)
        arrTAZ2 = depTAZ1 + s1s2_tt
        depTAZ2 = Max(arrTAZ2 + puTime, schEnd2 + puTime)   // Cannot leave before school ends
        s2h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ2, Dest: vecs.TAZID[kidIdx1], DepTime: depTAZ2})
        vecsOut.ID[i] = vecs.PersonID[kidIdx1]
        vecsOut.SchoolDep[i] = depTAZ1
        //vecsOut.ArrTAZ1[i] = arrTAZ2
        //vecsOut.DepTAZ1[i] = depTAZ2
        //vecsOut.TAZ1[i] = schTAZ2
        homeArr = depTAZ2 + s2h_tt
        vecsOut.HomeArr[i] = homeArr

        // Kid 2 record (2nd pickup)
        i = i + 1
        vecsOut.ID[i] = vecs.PersonID[kidIdx2]
        vecsOut.SchoolDep[i] = depTAZ2
        vecsOut.HomeArr[i] = homeArr

        // Adult record
        i = i + 1
        vecsOut.ID[i] = adultID
        vecsOut.KidID1[i] = vecs.PersonID[kidIdx1]
        vecsOut.KidID2[i] = vecs.PersonID[kidIdx2]
        vecsOut.WorkDep[i] = workDep
        vecsOut.ArrTAZ1[i] = arrTAZ1
        vecsOut.DepTAZ1[i] = depTAZ1
        vecsOut.TAZ1[i] = schTAZ1
        vecsOut.ArrTAZ2[i] = arrTAZ2
        vecsOut.DepTAZ2[i] = depTAZ2
        vecsOut.TAZ2[i] = schTAZ2
        vecsOut.HomeArr[i] = homeArr
        vecsOut.NPUDO[i] = 2

        nRecsAdded = 3
    end
    Return(nRecsAdded)
endMacro


/* 
    Macro to add separate PUDO tours
    The macro creates two new output tables (one for pickup tours and another for dropoff tours)
    Sometimes, two or three separate drop off tours will be needed if the kids being dropped off have different school start time periods
    Likewise, two or three separate pick up tours will be needed if the kids being picked up have different school end time periods
    Likewise if kids serviced by an adult share the same school end period, they are picked up in one tour
*/
Macro "Add Separate PUDO tours"(spec)
    RunMacro("Add Separate Dropoff Tours", spec)
    RunMacro("Add Separate Pickup Tours", spec)
endMacro


Macro "Add Separate Dropoff Tours"(spec)
    abm = spec.abmManager

    // Select kids who make carpool trips and need a separate tour from an adult
    filter = "DropoffTourFlag = 'S' and SchoolForwardModeCode = 2"
    abm.CreatePersonSet({Filter: filter, Activate: 1})

    // Loop over records (by DropoffPersonID and write one record for each person in a temporary table)
    pID = abm.[Person.DropoffPersonID]
    vwPUDO = RunMacro("Create Empty Tour File", {ViewName: "DOTours", NRecords: pID.Length}) // Allocate more than is necessary

    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "AddDropoffTours",
                MacroArgs: {SkimArgs: spec.SkimArgs},
                LoopOn: "DropoffPersonID",
                InputFields: {"PersonID", "DropoffPersonID", "Person.HouseholdID", "TAZID", "SchoolTAZ", "School_StartTime", "HomeToSchoolTime"},
                OutputView: vwPUDO,
                SortOrder: {{'DropoffPersonID', 'Ascending'}, {'School_StartTime', 'Ascending'}, {'Age', 'Ascending'}}
                }
    abm.Iterate(iterOpts)

    // Set basic vectors
    vecs = GetDataVectors(vwPUDO + "|", {"TourStartTime", "DestDepTime"}, {OptArray: 1})
    vTODF = RunMacro("Get TOD Vector", vecs.TourStartTime, spec.Periods)
    vTODR = RunMacro("Get TOD Vector", vecs.DestDepTime, spec.Periods)
    
    vecsSet = null
    vecsSet.TourType = Vector(pID.Length, "String", {Constant: "Mandatory"})
    vecsSet.TourPurpose = Vector(pID.Length, "String", {Constant: "Dropoff"})
    vecsSet.ForwardModeCode = Vector(pID.Length, "Short", {Constant: 2})
    vecsSet.ForwardMode = Vector(pID.Length, "String", {Constant: "Carpool"})
    vecsSet.ReturnModeCode = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.ReturnMode = Vector(pID.Length, "String", {Constant: "DriveAlone"})
    vecsSet.AssignForwardHalf = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.AssignReturnHalf = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.TODForward = vTODF
    vecsSet.TODReturn = vTODR
    SetDataVectors(vwPUDO + "|", vecsSet,)

    // Export to output file
    SetView(vwPUDO)
    n = SelectByQuery("NotEmpty", "several", "Select * where PerID <> null", )
    file = spec.DropoffTourFile
    ExportView(vwPUDO + "|NotEmpty", "FFB", file,,)
    CloseView(vwPUDO)
endMacro



Macro "Add Separate Pickup Tours"(spec)
    abm = spec.ABMManager

    // Select kids who make carpool trips and need a separate tour from an adult
    filter = "PickupTourFlag = 'S' and SchoolReturnModeCode = 2"
    abm.CreatePersonSet({Filter: filter, Activate: 1})

    // Loop over records (by DropoffPersonID and write one record for each person in a temporary table)
    pID = abm.[Person.PickupPersonID]
    vwPUDO = RunMacro("Create Empty Tour File", {ViewName: "PUTours", NRecords: pID.Length}) // Allocate more than is necessary

    schEndFld = CreateExpression(abm.PersonHHView, "School_EndTime", "School_StartTime + School_Duration",)
    
    mr = CreateObject("Model.Runtime")
    codeUI = mr.GetModelCodeUI()
    iterOpts = {UIName: codeUI,
                MacroName: "AddPickupTours",
                MacroArgs: {SkimArgs: spec.SkimArgs},
                LoopOn: "PickupPersonID",
                InputFields: {"PersonID", "PickupPersonID", "Person.HouseholdID", "TAZID", "SchoolTAZ", "School_EndTime"},
                OutputView: vwPUDO,
                SortOrder: {{'PickupPersonID', 'Ascending'}, {'School_EndTime', 'Ascending'}, {'Age', 'Ascending'}}
                }
    abm.Iterate(iterOpts)
    DestroyExpression(GetFieldFullSpec(abm.PersonHHView, schEndFld))

    // Set basic vectors
    vecs = GetDataVectors(vwPUDO + "|", {"TourStartTime", "DestDepTime"}, {OptArray: 1})
    vTODF = RunMacro("Get TOD Vector", vecs.TourStartTime, spec.Periods)
    vTODR = RunMacro("Get TOD Vector", vecs.DestDepTime, spec.Periods)

    vecsSet = null
    vecsSet.TourType = Vector(pID.Length, "String", {Constant: "Mandatory"})
    vecsSet.TourPurpose = Vector(pID.Length, "String", {Constant: "Pickup"})
    vecsSet.ForwardModeCode = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.ForwardMode = Vector(pID.Length, "String", {Constant: "DriveAlone"})
    vecsSet.ReturnModeCode = Vector(pID.Length, "Short", {Constant: 2})
    vecsSet.ReturnMode = Vector(pID.Length, "String", {Constant: "Carpool"})
    vecsSet.AssignForwardHalf = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.AssignReturnHalf = Vector(pID.Length, "Short", {Constant: 1})
    vecsSet.TODForward = vTODF
    vecsSet.TODReturn = vTODR
    SetDataVectors(vwPUDO + "|", vecsSet,)
    
    // Export to output file
    SetView(vwPUDO)
    n = SelectByQuery("NotEmpty", "several", "Select * where PerID <> null", )
    file = spec.PickupTourFile
    ExportView(vwPUDO + "|NotEmpty", "FFB", file,,)
    CloseView(vwPUDO)
endMacro



Macro "AddDropoffTours"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    recIndex = spec.RecordIndex
    SkimArgs = spec.MacroArgs.SkimArgs
    doffTime = 3
    tourBuffer = 15
    nDO = (endIdx - startIdx) + 1
    dim tourStart[2]

    // Get initial start time
    nTours = 1
    schTAZ1 = vecs.SchoolTAZ[startIdx]
    schStart1 = vecs.School_StartTime[startIdx]
    hs1_tt = vecs.HomeToSchoolTime[startIdx]
    tourStart[1] = schStart1 - hs1_tt
    if nDO > 1 then do // Two kids. Can combine or may need another tour. Depends on start times.
        schTAZ2 = vecs.SchoolTAZ[endIdx]
        schStart2 = vecs.School_StartTime[endIdx]
        hs2_tt = vecs.HomeToSchoolTime[endIdx]
        
        if schStart2 > (schStart1 + hs1_tt + hs2_tt + doffTime + tourBuffer) then do // More time than tour buffer left even if person went back home and then to the second school.
            nTours = 2
            tourStart[2] = schStart2 - hs2_tt
        end
        else do
            // Only one tour. But check if the tour needs to start early
            s1s2_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ1, Dest: schTAZ2, DepTime: schStart1 + doffTime})
            if (schStart2 - schStart1) < s1s2_tt + doffTime then do // Need to make up time by leaving home even earlier. First kid will be early to school.
                deltaTT = s1s2_tt - (schStart2 - schStart1) + doffTime
                tourStart[1] = schStart1 - hs1_tt - deltaTT
            end
        end
    end
    
    // Number of tours and tour start times determined. Write down records.
    vHHID = vecs.("Person.HouseholdID")
    for i = 1 to nTours do
        kidIdx = startIdx + (i-1)
        schStartTime = vecs.School_StartTime[kidIdx]

        vecsOut.HID[recIndex] = vHHID[kidIdx]
        vecsOut.HTAZ[recIndex] = vecs.TAZID[kidIdx]
        vecsOut.PerID[recIndex] = vecs.DropoffPersonID[kidIdx]
        vecsOut.Origin[recIndex] = vecs.TAZID[kidIdx]
        vecsOut.Destination[recIndex] = vecs.SchoolTAZ[kidIdx]
        vecsOut.TourStartTime[recIndex] = tourStart[i]
        vecsOut.DestArrTime[recIndex] = schStartTime
        vecsOut.DestDepTime[recIndex] = vecsOut.DestArrTime[recIndex] + doffTime
        vecsOut.ActivityStartTime[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.ActivityEndTime[recIndex] = vecsOut.DestDepTime[recIndex]
        s2h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.SchoolTAZ[kidIdx], Dest: vecs.TAZID[kidIdx], DepTime: vecsOut.DestDepTime[recIndex]})
        vecsOut.TourEndTime[recIndex] = vecsOut.DestDepTime[recIndex] + s2h_tt
        vecsOut.DropoffKidID1[recIndex] = vecs.PersonID[kidIdx]
        vecsOut.DropoffTAZ1[recIndex] = vecs.SchoolTAZ[kidIdx]
        vecsOut.ArrDropoff1[recIndex] = schStartTime
        vecsOut.DepDropoff1[recIndex] = schStartTime + doffTime
        vecsOut.NumDropoffs[recIndex] = 1
        recIndex = recIndex + 1
    end

    // Modify record if there are two kids dropped off on the same tour
    if nTours = 1 and nDO > 1 then do
        recIndex = recIndex - 1
        kidIdx1 = startIdx
        kidIdx2 = endIdx
        vecsOut.DestArrTime[recIndex] = tourStart[1] + hs1_tt
        vecsOut.ArrDropoff1[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.DepDropoff1[recIndex] = vecsOut.DestArrTime[recIndex] + doffTime
        vecsOut.DropoffKidID2[recIndex] = vecs.PersonID[kidIdx2]
        vecsOut.DropoffTAZ2[recIndex] = vecs.SchoolTAZ[kidIdx2]
        vecsOut.ArrDropoff2[recIndex] = vecsOut.DepDropoff1[recIndex] + s1s2_tt
        vecsOut.DepDropoff2[recIndex] = vecsOut.ArrDropoff2[recIndex] + doffTime
        vecsOut.DestDepTime[recIndex] = vecsOut.DepDropoff2[recIndex]
        vecsOut.ActivityStartTime[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.ActivityEndTime[recIndex] = vecsOut.DestDepTime[recIndex]
        s2h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.SchoolTAZ[kidIdx2], Dest: vecs.TAZID[kidIdx2], DepTime: vecsOut.DepDropoff2[recIndex]})
        vecsOut.TourEndTime[recIndex] = vecsOut.DepDropoff2[recIndex] + s2h_tt
        vecsOut.NumDropoffs[recIndex] = vecsOut.NumDropoffs[recIndex] + 1
    end

    Return(nTours)
endMacro


Macro "AddPickupTours"(spec)
    vecs = spec.InputVecs
    vecsOut = spec.OutputVecs
    startIdx = spec.StartIndex
    endIdx = spec.EndIndex
    recIndex = spec.RecordIndex
    SkimArgs = spec.MacroArgs.SkimArgs
    pupTime = 3
    tourBuffer = 15
    nPU = (endIdx - startIdx) + 1
    dim tourStart[2]

    // Get initial start time
    nTours = 1
    schTAZ1 = vecs.SchoolTAZ[startIdx]
    schEnd1 = vecs.School_EndTime[startIdx]
    hs1_tt_est = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[startIdx], Dest: schTAZ1, DepTime: schEnd1})
    hs1_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[startIdx], Dest: schTAZ1, DepTime: schEnd1 - hs1_tt_est})
    tourStart[1] = schEnd1 - hs1_tt
    if nPU > 1 then do // Two kids. Can combine or may need another tour. Depends on school end times.
        schTAZ2 = vecs.SchoolTAZ[endIdx]
        schEnd2 = vecs.School_EndTime[endIdx]
        hs2_tt_est = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[endIdx], Dest: schTAZ1, DepTime: schEnd2})
        hs2_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.TAZID[endIdx], Dest: schTAZ1, DepTime: schEnd2 - hs2_tt_est})
        
        if schEnd2 > (schEnd1 + hs1_tt + hs2_tt + pupTime + tourBuffer) then do // More time than tour buffer left even if person went back home and then to the second school.
            nTours = 2
            tourStart[2] = schEnd2 - hs2_tt
        end
        else do
            // Only one tour. But check if the tour needs to start early
            s1s2_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: schTAZ1, Dest: schTAZ2, DepTime: schEnd1 + pupTime})
            if (schEnd2 - schEnd1) < s1s2_tt + pupTime then do // Need to make up time by leaving home even earlier. First kid will be early to school.
                deltaTT = s1s2_tt - (schEnd2 - schEnd1) + pupTime
                tourStart[1] = schEnd1 - hs1_tt - deltaTT
            end
        end
    end
    
    // Number of tours and tour start times determined. Write down records.
    vHHID = vecs.("Person.HouseholdID")
    for i = 1 to nTours do
        kidIdx = startIdx + (i-1)
        schEndTime = vecs.School_EndTime[kidIdx]

        vecsOut.HID[recIndex] = vHHID[kidIdx]
        vecsOut.HTAZ[recIndex] = vecs.TAZID[kidIdx]
        vecsOut.PerID[recIndex] = vecs.PickupPersonID[kidIdx]
        vecsOut.Origin[recIndex] = vecs.TAZID[kidIdx]
        vecsOut.Destination[recIndex] = vecs.SchoolTAZ[kidIdx]
        vecsOut.TourStartTime[recIndex] = tourStart[i]
        vecsOut.DestArrTime[recIndex] = schEndTime
        vecsOut.DestDepTime[recIndex] = vecsOut.DestArrTime[recIndex] + pupTime
        vecsOut.ActivityStartTime[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.ActivityEndTime[recIndex] = vecsOut.DestDepTime[recIndex]
        s2h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.SchoolTAZ[kidIdx], Dest: vecs.TAZID[kidIdx], DepTime: vecsOut.DestDepTime[recIndex]})
        vecsOut.TourEndTime[recIndex] = vecsOut.DestDepTime[recIndex] + s2h_tt
        vecsOut.PickupKidID1[recIndex] = vecs.PersonID[kidIdx]
        vecsOut.PickupTAZ1[recIndex] = vecs.SchoolTAZ[kidIdx]
        vecsOut.ArrPickup1[recIndex] = schEndTime
        vecsOut.DepPickup1[recIndex] = schEndTime + pupTime
        vecsOut.NumPickups[recIndex] = 1
        recIndex = recIndex + 1
    end

    // Modify record if there are two kids picked up on the same tour
    if nTours = 1 and nPU > 1 then do
        recIndex = recIndex - 1
        kidIdx1 = startIdx
        kidIdx2 = endIdx
        vecsOut.DestArrTime[recIndex] = tourStart[1] + hs1_tt
        vecsOut.ArrPickup1[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.DepPickup1[recIndex] = vecsOut.DestArrTime[recIndex] + pupTime
        vecsOut.PickupKidID2[recIndex] = vecs.PersonID[kidIdx2]
        vecsOut.PickupTAZ2[recIndex] = vecs.SchoolTAZ[kidIdx2]
        vecsOut.ArrPickup2[recIndex] = vecsOut.DepPickup1[recIndex] + s1s2_tt
        vecsOut.DepPickup2[recIndex] = vecsOut.ArrPickup2[recIndex] + pupTime
        vecsOut.DestDepTime[recIndex] = vecsOut.DepPickup2[recIndex]
        vecsOut.ActivityStartTime[recIndex] = vecsOut.DestArrTime[recIndex]
        vecsOut.ActivityEndTime[recIndex] = vecsOut.DestDepTime[recIndex]
        s2h_tt = RunMacro("Get Auto TT", SkimArgs, {Orig: vecs.SchoolTAZ[kidIdx2], Dest: vecs.TAZID[kidIdx2], DepTime: vecsOut.DepPickup2[recIndex]})
        vecsOut.TourEndTime[recIndex] = vecsOut.DepPickup2[recIndex] + s2h_tt
        vecsOut.NumPickups[recIndex] = vecsOut.NumPickups[recIndex] + 1
    end

    Return(nTours)
endMacro
