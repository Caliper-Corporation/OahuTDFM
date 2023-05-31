Macro "PopulationSynthesis" (Args)

    // Check fields
    RunMacro("Prepare Seed Data", Args)

    // Compute block group marginals after DatabaseUSA correction
    DBUSA_HHSize6P = {Type: 'Norm', Mean: 6.48, StdDev: 0.23, MinValue: 6, MaxValue: 8}
    mSpec = {BGFile: Args.CensusBlockGroups, Filter: "InPeoria = 1", "DistributionHHSize6P": DBUSA_HHSize6P, OutputFile: Args.BGMarginals}
    RunMacro("Create Marginals File", mSpec)

    // Compute HH marginals at the block level (using block group marginals)
    nmSpec = {BlkFile: Args.CensusBlocks, Filter: "InPeoria = 1", MarginalsFile: Args.BGMarginals, OutputFile: Args.BlockMarginals}
    RunMacro("Create Nested Marginals File", nmSpec)

    // Run synthesis and perform tabulations
    spec = null
    spec.HHSeed = {FileName: Args.HHSeed, Filter: "NP > 0 and WGTP > 0", ID: "HHID", MatchingID: "PUMA", WeightField: "WGTP"}
    spec.PopSeed = {FileName: Args.PersonSeed, ID: "PersonID", HHID: "HHID"}
    spec.BGMarginals = {FileName: Args.BGMarginals, ID: "ID", MatchingID: "PUMA"}                           // Use in IPU (and not in IPF)
    spec.BlockMarginals = {FileName: Args.BlockMarginals, ID: "ID", MatchingID: "PUMAID", IPUID: "BG_ID"}   // Use in IPF
    spec.HHFile = Args.Households
    spec.PersonFile = Args.Persons
    spec.Tabulations = Args.BGTabulations
    spec.OutputFolder = Args.[Output Folder] + "\\Population\\"
    RunMacro("PopulationSynth", spec)

    // Run synthesis for dorm residents
    RunMacro("Dorm Residents Synthesis", Args)

    // Other post process on synthesis outputs (using the ABM Manager)
    spec.abmManager = Args.[ABM Manager]
    RunMacro("PopSyn Postprocess", spec)

    Return(true)
endMacro

/*
    Check HH and Person seed tables for any missing fields that are required by the synthesis
*/
Macro "Prepare Seed Data"(Args)
    reqHHSeedFields = {"HHID", "PUMA", "WGTP", "NP", "HINCP", "AdjInc", "R18", "R65", "Veh"}
    RunMacro("Check Fields", {File: Args.PUMS_Households, Fields: reqHHSeedFields})
    
    reqPersonSeedFields = {"PersonID", "HHID", "Sex", "AgeP", "INDP", "COW"}
    RunMacro("Check Fields", {File: Args.PUMS_Persons, Fields: reqPersonSeedFields})

    reqWrkIndFields = {"INDP", "INDP_Code", "Category"}
    RunMacro("Check Fields", {File: Args.PUMS_WorkIndustryCodes, Fields: reqWrkIndFields})

    // Export valid HH and Person seed records
    hhObj = CreateObject("Table", Args.PUMS_Households)
    perObj = CreateObject("Table", Args.PUMS_Persons)
    n = hhObj.SelectByQuery({SetName: "_Valid", Query: "NP > 0 and NP < 9 and WGTP > 0"})
    if n = 0 then
        Throw("No Valid records in Household Seed Data")
    hhSeedObj = hhObj.Export({FileName: Args.HHSeed, FieldNames: reqHHSeedFields})
    perSeedObj = perObj.Export({FileName: Args.PersonSeed, FieldNames: reqPersonSeedFields})
    hhObj = null
    perObj = null

    // Add Income, Seniors and Kids category fields to HH seed data
    flds = {{FieldName: "Inc_Category", Type: "Short", Description: "Income Category|1. < 25K|2. [25 50)|3. [50, 75)|4. [75, 100)|5. [100, 150)|6. [150, 200)|7. >= 200k"},
            {FieldName: "Kids_Category", Type: "Short", Description: "Kids Category|1. No Kids (age < 18) in HH|2. Kids present in HH"},
            {FieldName: "Seniors_Category", Type: "Short", Description: "Seniors Category|1. No Seniors (age >= 65) in HH|2. Seniors present in HH"}}
    hhSeedObj.AddFields({Fields: flds})
    vecs = hhSeedObj.GetDataVectors({FieldNames: {"HINCP", "ADJINC", "R18", "R65"}})
    v_inc = (vecs.ADJINC/1000000) * vecs.HINCP
    vecsSet = null
    vecsSet.Inc_Category = if v_inc < 25000 then 1
                            else if v_inc < 50000 then 2
                            else if v_inc < 75000 then 3
                            else if v_inc < 100000 then 4
                            else if v_inc < 150000 then 5
                            else if v_inc < 200000 then 6
                            else if v_inc >= 200000 then 7
                            else 3
    vecsSet.Seniors_Category = if vecs.R65 = 0 then 1 else 2
    vecsSet.Kids_Category = if vecs.R18 = 0 then 1 else 2
    hhSeedObj.SetDataVectors({FieldData: vecsSet})

    // Add Industry category field and fill in person seed. Need the industry codes table to do so
    flds = {{FieldName: "WorkIndustry", Type: "Short", Description: "Work Industry: (12 categories)"},
            {FieldName: "IndustryCategory", Type: "Short", Description: "Industry Category (10 categories) used in synthesis"}}
    perSeedObj.AddFields({Fields: flds})

    wrkCodesObj = CreateObject("Table", Args.PUMS_WorkIndustryCodes)
    jvObj = perSeedObj.Join({Table: wrkCodesObj, LeftFields: {"INDP"}, RightFields: {"INDP"}})
    vecs = jvObj.GetDataVectors({FieldNames: {"IndP_Code", "Category"}})
    vecsSet = null
    vecsSet.WorkIndustry = vecs.IndP_Code
    vecsSet.IndustryCategory = if vecs.Category = null then 10 else vecs.Category
    jvObj.SetDataVectors({FieldData: vecsSet})

    jvObj = null
    wrkCodesObj = null
    hhSeedObj = null
    perSeedObj = null
endMacro


/*
    Generic macro to check if the provided file contains the provided fields
*/
Macro "Check Fields"(spec)
    dm = CreateObject("DataManager")
    vw = dm.AddDataSource("File", {FileName: spec.File})
    {flds, specs} = GetFields(vw,)
    flds = flds.Map(do (f) Return(if f[1] = "[" then SubString(f, 2, StringLength(f) - 2) else f) end)
    for fld in spec.Fields do
        if flds.position(fld) = 0 then
            Throw(Printf("Field '%s' not found in file '%s'", {fld, spec.File}))
    end
endMacro


/*
    Macro that uses the Census Block Group data and creates HH and Person marginals for synthesis
    1. The HH by Size fields are corrected using the Database USA correction.
       The correction is based on the average HH Size value for 6PlusPersonHHs.
       This is a normal distribution with a mean of 6.48 and StdDev of 0.23
       The correction modifies the HH Size fields (and the total households)

    2. All other HH marginals are rescaled based on the new household totals

    3. Person marginals are appropiately collapsed into fewer categories

    4. An output marginals file is created
*/
Macro "Create Marginals File"(mSpec)
    // Check for presence of required fields
    mainFlds = {"ID", "PUMA", "InPeoria", "Households", "Population", "Population in GQ"}
    hhSizeFlds = {"HH_1Person", "HH_2Person", "HH_3Person", "HH_4Person", 
                  "HH_5Person", "HH_6Person", "HH_7+Person"}
    hhOthFlds = {"HH_0Veh", "HH_1Veh", "HH_2Veh", "HH_3+Veh", 
                "HH_No_Kids", "HH_With_Kids",
                "HH_No_Seniors", "HH_With_Seniors",   
                "HH_Inc_LT25", "HH_Inc_25to50", "HH_Inc_50to75", "HH_Inc_75to100",
                "HH_Inc_100to150", "HH_Inc_150to200", "HH_Inc_GT200"}
    popGenderFlds = {"Male", "Female"}
    popAgeFlds = { "Age <5", "Age 5 to 9", "Age 10 to 14", "Age 15 to 19", "Age 20 to 24", "Age 25 to 34",
                    "Age 35 to 44", "Age 45 to 54", "Age 55 to 59", "Age 60 to 64", "Age 65 to 74", "Age 75 to 84", "Age 85+"}
    popIndFlds = {"EC 16+_Ind: Ag/forestry/fish/mine", 
                    "EC 16+_Ind: Manufacturing", 
                    "EC 16+_Ind: Construction", "EC 16+_Ind: Transportation/warehouse", 
                    "EC 16+_Ind: Wholesale trade", 
                    "EC 16+_Ind: Retail trade", 
                    "EC 16+_Ind: Information",  "EC 16+_Ind: Finance/ins/RE/rental", "EC 16+_Ind: Prof/scientific/admin", 
                    "EC 16+_Ind: Ed/health/soc services", 
                    "EC 16+_Ind: Art/ent/rec/acc/food", "EC 16+_Ind: Other (ex public admin)", 
                    "EC 16+_Ind: Public administration"}
    flds = mainFlds + hhSizeFlds + hhOthFlds + popGenderFlds + popAgeFlds + popIndFlds
    RunMacro("Check Fields", {File: mSpec.BGFile, Fields: flds})
    
    // Export relevant fields to marginals output file
    bgObj = CreateObject("Table", mSpec.BGFile)
    if mSpec.Filter <> null then do
        setName = "Selection"
        n = bgObj.SelectByQuery({SetName: "Selection", Query: mSpec.Filter})
        if n = 0 then
            Throw(Printf("No valid records in block group data for filter: '%s'", {mSpec.Filter}))
    end
    //mObj = bgObj.Export({FileName: mSpec.OutputFile, FieldNames: flds})
    ExportView(bgObj.GetView() + "|" + setName, "FFB", mSpec.OutputFile, flds,)
    mObj = CreateObject("Table", mSpec.OutputFile)
    bgObj = null
    
    vecs = mObj.GetDataVectors({FieldNames: flds})
    nRecs = vecs.ID.Length

    // Generate a vector corresponding to the normal distrbution of the HHSize6Plus term
    SetRandomSeed(99991)
    params = {mu: mSpec.DistributionHHSize6P.Mean, sigma: mSpec.DistributionHHSize6P.StdDev}
    vNorm = RandSamples(nRecs, mSpec.DistributionHHSize6P.Type, params)
    
    // Apply bounds if supplied
    minVal = mSpec.DistributionHHSize6P.MinValue
    if minVal <> null then
        vNorm = if vNorm < minVal then minVal else vNorm

    maxVal = mSpec.DistributionHHSize6P.MaxValue
    if maxVal <> null then
        vNorm = if vNorm > maxVal then maxVal else vNorm

    // Add necessary fields
    fldsToAdd = {"HHSize1", "HHSize2", "HHSize3", "HHSize4", "HHSize5", "HHSize6P",
                 "Age_0_4", "Age_5_14", "Age_15_19", "Age_20_34", "Age_35_54", "Age_55_64", "Age_65P"}
    for i = 1 to 10 do
        fldsToAdd = fldsToAdd + {"Industry" + String(i)}
    end
    newFlds = fldsToAdd.Map(do (f) Return({FieldName: f, Type: "real", Width: 12, Decimals: 2}) end)
    mObj.AddFields({Fields: newFlds})

    // DB USA Adjustment
    vHH = vecs.Households
    vPop = vecs.Population - vecs.[Population in GQ]
    vCurrPop1to5 = vecs.HH_1Person + 2*vecs.HH_2Person + 3*vecs.HH_3Person + 4*vecs.HH_4Person + 5*vecs.HH_5Person
    vCurrHH6P = vecs.HH_6Person + vecs.[HH_7+Person]
    vNewPop6P = vCurrHH6P * vNorm
    vNewPop1to6 = vCurrPop1to5 + vNewPop6P
    vRatio = if vNewPop1to6 > 0 then vPop/vNewPop1to6 else 1
    
    vecsSet = null
    vecsSet.HHSize1 = vecs.HH_1Person * vRatio
    vecsSet.HHSize2 = vecs.HH_2Person * vRatio
    vecsSet.HHSize3 = vecs.HH_3Person * vRatio
    vecsSet.HHSize4 = vecs.HH_4Person * vRatio
    vecsSet.HHSize5 = vecs.HH_5Person * vRatio
    vecsSet.HHSize6P = vCurrHH6P * vRatio

    // Compute new HHs and ratio of new to old HHs
    vecsSet.Households = vecsSet.HHSize1 + vecsSet.HHSize2 + vecsSet.HHSize3 + vecsSet.HHSize4 + vecsSet.HHSize5 + vecsSet.HHSize6P
    vHHRatio = if vecs.Households > 0 then vecsSet.Households/vecs.Households else 1

    // Adjust all other marginals
    for fld in hhOthFlds do
        vecsSet.(fld) = vecs.(fld) * vHHRatio
    end

    // Collapse population age categories
    vPopRatio = if vecs.Population > 0 then (1 - nz(vecs.[Population in GQ])/nz(vecs.Population)) else 1
    vecsSet.Male = vecs.Male * vPopRatio
    vecsSet.Female = vecs.Female * vPopRatio

    vecsSet.[Age_0_4] = vecs.[Age <5] * vPopRatio
    vecsSet.[Age_5_14] = (vecs.[Age 5 to 9] + vecs.[Age 10 to 14]) * vPopRatio
    vecsSet.[Age_15_19] = vecs.[Age 15 to 19] * vPopRatio
    vecsSet.[Age_20_34] = (vecs.[Age 20 to 24] + vecs.[Age 25 to 34]) * vPopRatio
    vecsSet.[Age_35_54] = (vecs.[Age 35 to 44] + vecs.[Age 45 to 54]) * vPopRatio
    vecsSet.[Age_55_64] = (vecs.[Age 55 to 59] + vecs.[Age 60 to 64]) * vPopRatio
    vecsSet.[Age_65P] = (vecs.[Age 65 to 74] + vecs.[Age 75 to 84] + vecs.[Age 85+]) * vPopRatio

    // Pop Ind categories
    vecsSet.Industry1 = vecs.[EC 16+_Ind: Ag/forestry/fish/mine] * vPopRatio
    vecsSet.Industry2 = vecs.[EC 16+_Ind: Manufacturing] * vPopRatio
    vecsSet.Industry3 = (vecs.[EC 16+_Ind: Construction] + vecs.[EC 16+_Ind: Transportation/warehouse]) * vPopRatio
    vecsSet.Industry4 = vecs.[EC 16+_Ind: Wholesale trade] * vPopRatio
    vecsSet.Industry5 = vecs.[EC 16+_Ind: Retail trade] * vPopRatio
    vecsSet.Industry6 = (vecs.[EC 16+_Ind: Information] + vecs.[EC 16+_Ind: Finance/ins/RE/rental] + vecs.[EC 16+_Ind: Prof/scientific/admin]) * vPopRatio
    vecsSet.Industry7 = vecs.[EC 16+_Ind: Ed/health/soc services] * vPopRatio
    vecsSet.Industry8 = (vecs.[EC 16+_Ind: Art/ent/rec/acc/food] + vecs.[EC 16+_Ind: Other (ex public admin)]) * vPopRatio
    vecsSet.Industry9 = vecs.[EC 16+_Ind: Public administration] * vPopRatio
    
    vSum1to9 = vecsSet.Industry1
    for i = 2 to 9 do
        vSum1to9 = vSum1to9 + vecsSet.("Industry" + String(i))
    end
    vecsSet.Industry10 = vPop - vSum1to9

    // Set Vectors
    mObj.SetDataVectors({FieldData: vecsSet})

    // Drop old HH by Size, old Age fields and old occupation fields
    mObj.DropFields({FieldNames: hhSizeFlds + popAgeFlds + popIndFlds})
    mObj = null
endMacro


/*
    Macro that uses the Census Block data and Block Group data creates HH Size marginals at the block level
    1. The proportion of HHs in each block is calculated (within the block group) solely based on the block data

    2. The HH by Size fields for each block is assumed to be the same as the HH by Size distribution for the corresponding block group

    3. An output nested marginals file is created with HH by Size at the block level
*/
Macro "Create Nested Marginals File"(nmSpec)
    // Open the census block file and export relevant fields
    flds = {"ID", "BG_ID", "TAZ_ID", "InPeoria", "HU_Occupied"}
    RunMacro("Check Fields", {File: nmSpec.BlkFile, Fields: flds})
    
    blkObj = CreateObject("Table", nmSpec.BlkFile)
    if nmSpec.Filter <> null then do
        setName = "Selection"
        n = blkObj.SelectByQuery({SetName: "Selection", Query: nmSpec.Filter})
        if n = 0 then
            Throw(Printf("No valid records in block data for filter: '%s'", {nmSpec.Filter}))
    end
    //mObj = blkObj.Export({FileName: nmSpec.OutputFile, FieldNames: flds})
    ExportView(blkObj.GetView() + "|" + setName, "FFB", nmSpec.OutputFile, flds,)
    nmObj = CreateObject("Table", nmSpec.OutputFile)
    blkObj = null
    
    // Add empty output fields to nested marginals file. Retain same names for the fields used in the synthesis.
    marginalFlds  = {"HHSize1", "HHSize2", "HHSize3", "HHSize4", "HHSize5", "HHSize6P",
                    "HH_Inc_LT25", "HH_Inc_25to50", "HH_Inc_50to75", "HH_Inc_75to100", 
                    "HH_Inc_100to150", "HH_Inc_150to200", "HH_Inc_GT200", 
                    "HH_No_Kids", "HH_With_Kids", "HH_No_Seniors", "HH_With_Seniors",
                    "HH_0Veh", "HH_1Veh", "HH_2Veh", "HH_3+Veh"}
    newFlds = marginalFlds.Map(do (f) Return({FieldName: f, Type: "real", Width: 12, Decimals: 2}) end)
    newFlds = newFlds + {{FieldName: "BG_HU_Occupied", Type: "real", Width: 12, Decimals: 2},
                         {FieldName: "PUMAID", Type: "integer", Width: 12}}
    nmObj.AddFields({Fields: newFlds})

    // Aggregate current Block HHs to get BG HH totals.
    vwNM = nmObj.GetView()
    vwAgg = AggregateTable("AggBlks", vwNM + "|", "MEM", "AggBlk", "BG_ID", {{"HU_Occupied", "Sum",}},)
    {flds, specs} = GetFields(vwAgg,)
    vwJ = JoinViews("Blk_Agg", GetFieldFullSpec(vwNM, "BG_ID"), specs[1],)
    v = GetDataVector(vwJ + "|", specs[2],)
    SetDataVector(vwJ + "|", "BG_HU_Occupied", v,)
    CloseView(vwJ)
    CloseView(vwAgg)

    // Join Blk to BG to new table and process
    bgObj = CreateObject("Table", nmSpec.MarginalsFile)
    vwBG = bgObj.GetView()

    // Join Blk to BG to new table and process
    vwJ = JoinViews("Blk_BG", GetFieldFullSpec(vwNM, "BG_ID"), GetFieldFullSpec(vwBG, "ID"),)
    mFldSpecs = marginalFlds.Map(do (f) Return(GetFieldFullSpec(vwBG, f)) end)
    nmFldSpecs = marginalFlds.Map(do (f) Return(GetFieldFullSpec(vwNM, f)) end)
    
    flds = {"HU_Occupied", "BG_HU_Occupied", "PUMA"} + mFldSpecs
    vecs = GetDataVectors(vwJ + "|", flds, {OptArray:1})
    vRatio = if nz(vecs.BG_HU_Occupied) = 0 then 0 else vecs.HU_Occupied/vecs.BG_HU_Occupied
    
    vecsSet = null
    vecsSet.PUMAID = vecs.PUMA
    for i = 1 to marginalFlds.length do
        mFld = mFldSpecs[i]
        nmFld = nmFldSpecs[i]
        vecsSet.(nmFld) = vecs.(mFld) * vRatio
    end
    SetDataVectors(vwJ + "|", vecsSet,)
    CloseView(vwJ)
    
    mObj = null
    nmObj = null
endMacro


/*
    1. Run Nested Synthesis with IPU
    - IPF HH Marginals: HH by Size at Block level
                        HH by Vehicles at Block level
                        HH by Income Block level
                        HH by presence of kids Block level
                        HH by presence of seniors Block level
    
    - IPU HH Marginals: All IPF HH marginals at the BG level

    - IPU Person Marginals: Persons by Gender at BG level
                            Persons by Age at BG level
                            Persons by WorkIndustry at BG level

    2. Create one-way tabulations from synthesized files
*/
Macro "PopulationSynth"(spec)
    o = CreateObject("PopulationSynthesis")
    o.HouseholdFile(spec.HHSeed)
    o.PersonFile(spec.PopSeed)
    o.MarginalFile(spec.BlockMarginals)     
    o.IPUMarginalFile(spec.BGMarginals)            

    HHDimSynmargData1 = {
        { Name: "HHSize1" , Value: {1, 2}}, 
        { Name: "HHSize2" , Value: {2, 3}}, 
        { Name: "HHSize3" , Value: {3, 4}}, 
        { Name: "HHSize4" , Value: {4, 5}}, 
        { Name: "HHSize5" , Value: {5, 6}},
        { Name: "HHSize6P" , Value: {6, 99}}
    }
    o.AddHHMarginal({ Field: "NP", Value: HHDimSynmargData1  , NewFieldName: "HHSize"})
    
    HHDimSynmargData2 = {
        { Name: "HH_0Veh" , Value: {0, 1}}, 
        { Name: "HH_1Veh" , Value: {1, 2}}, 
        { Name: "HH_2Veh" , Value: {2, 3}}, 
        { Name: "HH_3+Veh" , Value: {3, 99}}
    }
    o.AddHHMarginal({ Field: "Veh", Value: HHDimSynmargData2  , NewFieldName: "Vehicles"})
    
    HHDimSynmargData3 = {
        { Name: "HH_Inc_LT25" , Value: {1, 2}}, 
        { Name: "HH_Inc_25to50" , Value: {2, 3}}, 
        { Name: "HH_Inc_50to75" , Value: {3, 4}}, 
        { Name: "HH_Inc_75to100" , Value: {4, 5}}, 
        { Name: "HH_Inc_100to150" , Value: {5, 6}}, 
        { Name: "HH_Inc_150to200" , Value: {6, 7}}, 
        { Name: "HH_Inc_GT200" , Value: {7, 99}}
    }
    o.AddHHMarginal({ Field: "Inc_Category", Value: HHDimSynmargData3  , NewFieldName: "Inc_Category"})
    
    HHDimSynmargData4 = {
        { Name: "HH_No_Kids" , Value: {1, 2}},
        { Name: "HH_With_Kids" , Value: {2, 99}} 
    }
    o.AddHHMarginal({ Field: "Kids_Category", Value: HHDimSynmargData4  , NewFieldName: "Kids_Category"})
    
    HHDimSynmargData5 = {
        { Name: "HH_No_Seniors" , Value: {1, 2}},
        { Name: "HH_With_Seniors" , Value: {2, 99}}
    }
    o.AddHHMarginal({ Field: "Seniors_Category", Value: HHDimSynmargData5  , NewFieldName: "Seniors_Category"})
    
    HHDimIPUmargData1 = HHDimSynmargData1
    o.AddIPUHouseholdMarginal({ Field: "NP", Value: HHDimIPUmargData1  , NewFieldName: "HHSize"})
    
    HHDimIPUmargData2 = HHDimSynmargData2
    o.AddIPUHouseholdMarginal({ Field: "Veh", Value: HHDimIPUmargData2  , NewFieldName: "Vehicles"})
    
    HHDimIPUmargData3 = HHDimSynmargData3
    o.AddIPUHouseholdMarginal({ Field: "Inc_Category", Value: HHDimIPUmargData3  , NewFieldName: "Inc_Category"})
    
    HHDimIPUmargData4 = HHDimSynmargData4
    o.AddIPUHouseholdMarginal({ Field: "Kids_Category", Value: HHDimIPUmargData4  , NewFieldName: "Kids_Category"})
    
    HHDimIPUmargData5 = HHDimSynmargData5
    o.AddIPUHouseholdMarginal({ Field: "Seniors_Category", Value: HHDimIPUmargData5  , NewFieldName: "Seniors_Category"})

    PersonDimIPUmargData1 = {
        { Name: "Male" , Value: {1, 2}}, 
        { Name: "Female" , Value: {2, 99}}
    }
    o.AddIPUPersonMarginal({ Field: "SEX", Value: PersonDimIPUmargData1  , NewFieldName: "Gender"})

    PersonDimIPUmargData2 = {
        { Name: "Age_0_4" , Value: {0, 5}}, 
        { Name: "Age_5_14" , Value: {5, 15}}, 
        { Name: "Age_15_19" , Value: {15, 20}}, 
        { Name: "Age_20_34" , Value: {20, 35}}, 
        { Name: "Age_35_54" , Value: {35, 55}}, 
        { Name: "Age_55_64" , Value: {55, 65}}, 
        { Name: "Age_65P" , Value: {65, 999}}
    }
    o.AddIPUPersonMarginal({ Field: "AGEP", Value: PersonDimIPUmargData2  , NewFieldName: "Age"})

    PersonDimIPUmargData3 = {
        { Name: "Industry1" , Value: {1, 2}}, 
        { Name: "Industry2" , Value: {2, 3}}, 
        { Name: "Industry3" , Value: {3, 4}}, 
        { Name: "Industry4" , Value: {4, 5}}, 
        { Name: "Industry5" , Value: {5, 6}}, 
        { Name: "Industry6" , Value: {6, 7}}, 
        { Name: "Industry7" , Value: {7, 8}}, 
        { Name: "Industry8" , Value: {8, 9}}, 
        { Name: "Industry9" , Value: {9, 10}}, 
        { Name: "Industry10" , Value: {10, 99}}
    }
    o.AddIPUPersonMarginal({ Field: "IndustryCategory", Value: PersonDimIPUmargData3  , NewFieldName: "IndustryCategory"})

    o.OutputHouseholdsFile   = spec.HHFile
    o.OutputPersonsFile      = spec.PersonFile
    o.IPUIncidenceOutputFile = spec.OutputFolder + "Intermediate\\IPUIncidence.bin"
    o.ExportIPUWeights(spec.OutputFolder + "Intermediate\\IPUWeights")
    o.ReportExtraPersonsField("WorkIndustry", "WorkIndustry")
    o.ReportExtraPersonsField("COW", "ClassOfWorker")
    o.MarginError = 0.01
    ok = o.Run()
    //if ok then 
    //    res = o.GetResults({CalculateSummaryStatistics: true, FileName: outputFolder + "PopSynth_Statistics.bin", OpenResults: 0})

    RunMacro("Add PopSyn TAZ and BG Info", spec)
    
    // Call macro for tabulations
    opts = null
    opts.HHFile = spec.HHFile
    opts.PersonFile = spec.PersonFile
    opts.OutputFile = spec.Tabulations
    opts.GroupBy = "BGID"
    
    opts.HHTabulations.HHSize = HHDimIPUmargData1
    opts.HHTabulations.Vehicles = HHDimIPUmargData2
    opts.HHTabulations.Inc_Category = HHDimIPUmargData3
    opts.HHTabulations.Kids_Category = HHDimIPUmargData4
    opts.HHTabulations.Seniors_Category = HHDimIPUmargData5
    
    opts.PersonTabulations.Gender = PersonDimIPUmargData1
    opts.PersonTabulations.Age = PersonDimIPUmargData2
    opts.PersonTabulations.IndustryCategory = PersonDimIPUmargData3
    RunMacro("Generate Tabulations", opts)
endmacro


/*
    - Add Block group and TAZ ID to output HH pop synth database
    - Note that the file contains ZoneID, which is the BlockID and renamed as such
*/
Macro "Add PopSyn TAZ and BG Info"(spec)
    // Add fields to store BG and TAZ ID
    hhObj = CreateObject("Table", spec.HHFile)
    newFlds = newFlds + {{FieldName: "TAZID", Type: "integer", Width: 12, Description: "TAZID field from the TAZ layer"},
                         {FieldName: "BGID", Type: "integer", Width: 12, Description: "Block Group ID"}}
    hhObj.AddFields({Fields: newFlds})

    // Rename ZoneID field to BlockID
    hhObj.RenameField({FieldName: "ZoneID", NewName: "BlockID"})

    // Transfer block group ID field to HH database
    blkObj = CreateObject("Table", spec.BlockMarginals.FileName)
    jvObj = hhObj.Join({Table: blkObj, LeftFields: {"BlockID"}, RightFields: {"ID"}})
    jvObj.BGID = jvObj.BG_ID
    jvObj.TAZID = jvObj.TAZ_ID
    jvObj = null
endMacro


/*
    Generic model to perform tabulations after synthesis
    Uses the same input format as pop synth
    Can ideally be integrated into the procedure in future TC versions
*/
Macro "Generate Tabulations"(opts)
    dm = CreateObject("DataManager")
    vwHH = dm.AddDataSource("HH", {FileName: opts.HHFile})
    vwPer = dm.AddDataSource("Per", {FileName: opts.PersonFile})
    vwPerHH = JoinViews("PersonHH", GetFieldFullSpec(vwPer, "HouseholdID"), GetFieldFullSpec(vwHH, "HouseholdID"), )
    if opts.OutputFile = null or TypeOf(opts.OutputFile) <> "string" then
        Throw("Please specify valid 'OutputFile' option for population synthesis output tabulations")

    groupByFld = opts.GroupBy
    
    // HH Tabulations
    specs = {View: vwHH, Tabulations: opts.HHTabulations}
    exprsHH = RunMacro("Create Expressions for Tabulations", specs)

    specs = {View: vwHH, ExpressionNames: exprsHH, GroupBy: groupByFld}
    vwHHTab = RunMacro("Write Tabulations", specs)

    // Person Tabulations
    specs = {View: vwPerHH, Tabulations: opts.PersonTabulations}
    exprsPer = RunMacro("Create Expressions for Tabulations", specs)

    specs = {View: vwPerHH, ExpressionNames: exprsPer, GroupBy: groupByFld}
    vwPerTab = RunMacro("Write Tabulations", specs)
    
    CloseView(vwPerHH)
    dm = null

    // Join HH and Person tabulations and export to final table
    vwJ = JoinViews("HHPerTab", GetFieldFullSpec(vwHHTab, groupByFld), GetFieldFullSpec(vwPerTab, groupByFld),)
    fldsToExport = {GetFieldFullSpec(vwHHTab, groupByFld)} + exprsHH + exprsPer
    ExportView(vwJ + "|", "FFB", opts.OutputFile, fldsToExport,)
    CloseView(vwJ)
    CloseView(vwHHTab)
    CloseView(vwPerTab)
endMacro


/*
    Generate expressions that are aggragated to produce the synthesis tabulations
*/
Macro "Create Expressions for Tabulations"(specs)
    vw = specs.View
    exprNames = null
    tabs = specs.Tabulations
    for i = 1 to tabs.length do
        outFld = tabs[i][1]
        vals = tabs[i][2]
        for val in vals do
            fldName = val.Name
            if fldName[1] = "[" then // Remove first and last bracket
                fldName = SubString(fldName, 2, StringLength(fldName) - 2)
            
            exprName = fldName + "_Out"
            range = val.Value
            if TypeOf(range) = "array" then
                expr = printf("if %s >= %lu and %s < %lu then 1 else 0", {outFld, range[1], outFld, range[2]})
            else
                expr = printf("if %s = %lu then 1 else 0", {outFld, range})

            newExpr = CreateExpression(vw, exprName, expr,)
            exprNames = exprNames + {newExpr}
        end
    end
    Return(exprNames)
endMacro


/*
    Aggregate expressions using the 'GroupBy' field to produce an In-Memory aggregation table
*/
Macro "Write Tabulations"(specs)
    exprNames = specs.ExpressionNames
    aggFlds = exprNames.Map(do (f) Return({f, "sum",}) end)
    vw_agg = AggregateTable("Aggregations", specs.View + "|", "MEM", "Agg", specs.GroupBy, aggFlds, null)
    for exprName in exprNames do
        DestroyExpression(GetFieldFullSpec(specs.View, exprName))
    end
    Return(vw_agg)
endMacro


/*
    Synthesis post process that fills several HH fields based on person level aggregations
    Adds metadata to the table (balloon help)
*/
Macro "PopSyn Postprocess"(spec)
    abm = spec.abmManager
    hhFile = spec.HHFile
    personFile = spec.PersonFile
    
    // Add fields to HH database
    newFlds = { {Name: "Adults", Type: "Short", Description: "Number of adults in the household (Age >= 18)"},
                {Name: "Kids", Type: "Short", Description: "Number of kids in the household (Age < 18)"},
                {Name: "PreSchKids", Type: "Short", Description: "Number of pre-school kids in the household (Age < 5)"},
                {Name: "Females", Type: "Short"},
                {Name: "Males", Type: "Short"},
                {Name: "Workers", Type: "Short", Description: "Number of workers in the household (IndustryCategory < 10)"},
                {Name: "Seniors", Type: "Short", Description: "Number of seniors in the household (Age >= 65)"},
                {Name: "IncomeLevel", Type: "Short", Description: "HH Income level: 1 if HHIncome is below $50K, 2 if HHIncome is [50K, 100K), 3 if HHIncome is gte $100K"},
                {Name: "AvgWrkIncCategory", Type: "Short", Description: "Avg Worker Income level: 1 if HHIncome/HHWorkers is below $50K, 2 if HHIncome/HHWorkers is [50K, 100K), 3 if HHIncome/HHWorkers is gte $100K"}
               }
    abm.AddHHFields(newFlds)
    
    // Aggregate person fields and fill in household table
    aggOpts.Spec = {{Kids: "(Age < 18).Count", DefaultValue: 0},   // (Condition).Count
                    {PreSchKids: "(Age < 5).Count", DefaultValue: 0},
                    {Adults: "(Age >= 18).Count", DefaultValue: 0},
                    {Seniors: "(Age >= 65).Count", DefaultValue: 0},
                    {Males: "(Gender = 1).Count", DefaultValue: 0},
                    {Females: "(Gender = 2).Count", DefaultValue: 0},
                    {Workers: "(IndustryCategory < 10).Count", DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)

    // Fill IncomeLevel and AvgWrkIncCategory fields
    vecs = abm.GetHHVectors({"Inc_Category", "Workers"})
    vecsSet = null
    vecsSet.IncomeLevel = if vecs.Inc_Category <= 2 then 1 
                            else if vecs.Inc_Category <= 4 then 2 
                            else 3
    vecsSet.AvgWrkIncCategory = if vecs.Inc_Category <= 2 then 1
                                else if vecs.Inc_Category <= 4 and vecs.Workers >= 2 then 1
                                else if vecs.Inc_Category <= 4 and vecs.Workers < 2 then 2
                                else if vecs.Inc_Category = 5 and vecs.Workers >= 3 then 1
                                else if vecs.Inc_Category = 5 and vecs.Workers = 2 then 2
                                else if vecs.Inc_Category = 5 and vecs.Workers < 2 then 3
                                else if vecs.Inc_Category >= 6 and vecs.Workers >= 4 then 1
                                else if vecs.Inc_Category >= 6 and vecs.Workers = 3 then 2
                                else if vecs.Inc_Category >= 6 and vecs.Workers < 3 then 3
                                else null
    abm.SetHHVectors(vecsSet)

    // Add pop syn metadata
    opt = {HHView: abm.HHView, PersonView: abm.PersonView}
    RunMacro("Add PopSyn Metadata", opt)

    // Export data
    abm.ExportHHView({File: hhFile})
    abm.ExportPersonView({File: personFile})
endMacro


/*
    Add metadata (balloon help) to certain fields in the Person and HH output databases
*/
Macro "Add PopSyn Metadata"(spec)
    vwHH = spec.HHView
    vwP = spec.PersonView

    // Add HH metadata
    items = {ZoneID: "Census Block Group layer ID field",
             SubzoneID: "Census Block layer ID field",
             "Veh_Category": "1: Zero Vehicles|2: One Vehicle|3. Two Vehicles|4: Three or more vehicles",
             "Inc_Category": "HH Income|1. Less than 25k|2. 25k to 50k|3. 50k to 75k|4. 75k to 100k|5. 100k to 150k|6. 150k to 200k|7. More than 200K",
             "Kids_Category": "1. HH without kids (Age < 18)|2. HH with kids",
             "Seniors_Category": "1. HH without seniors (Age >= 65)|2. HH with seniors"
            }
    RunMacro("Add Balloon Help", vwHH, items)

    // Add Person metadata
    indCatStr = "Worker industry category|" + 
                "1. Agriculture/Mining|2. Manufacturing|3. Utilities, Construction, Transportation, Waste Management|" +
                "4. Wholesale Trade| 5. Retail Trade|" +   
                "6. Information, Finance, Insurance, Professional, Scientific, Technical, Management, Real Estate|" +
                "7. Education/Health/Social Services|8. Food/Other services|" +
                "9. Public Administration and Military|10. Not in Labor Force or unemployed"
    
    indStr = "Tagged field with worker industry using NAICS classification codes|" + 
             "1. Agriculture|2. Manufacturing and Mining|3. Utilities, Construction, Transportation, Waste Management|" +
             "4. Wholesale Trade|5. Retail Trade|" +   
             "6. Information, Finance, Insurance, Professional, Scientific, Technical, Management, Real Estate|" +
             "7. Education|8. Health Care|9. Arts, Entertainment, Food and Other Services|" +
             "10. Public Administration|11. Military|12. Not in Labor Force or unemployed"

    cowStr = "Class of Worker: Tagged|1. Employee of a private for-profit company, or of an individual for wages|" +
             "2. Employee of a private not-for-profit organization|" +
             "3. Employee of local government|4. Employee of state government|5. Employee of federal goverment|" + 
             "6. Self-employed in own not incorporated business|" +
             "7. Self-employed in own incorporated business|" +
             "8. Working without pay in family business|" +
             "9. Unemployed or never worked|" +
             "Blank. Less than 16 years old or not in labor force for more than 5 years"
    
    items = {Gender: "1. Male|2. Female",
             IndustryCategory: indCatStr,
             WorkIndustry: indStr,
             ClassOfWorker: cowStr
            }
    RunMacro("Add Balloon Help", vwP, items)
    dm = null
endMacro


/*
    Add balloon help given the view and an option array pair of field names and descriptions
*/
Macro "Add Balloon Help"(vw, items)
    strct = GetTableStructure(vw, {"Include Original" : "True"})
    names = strct.Map(do (f) Return(f[1]) end)
    for item in items do
        pos = names.position(item[1])
        if pos > 0 then
            strct[pos][8] = item[2]
    end
    ModifyTable(vw, strct)
endMacro


/* Dorm Synthesis
    Synthesize HHs and Persons for University residents from University Group Quarters information
    - HH TAZ and Univ TAZ known
    - HH Size = 1 (Each person constitutes a HH)
    - Income distribution (low or low-medium) based on input parameter
    - Age distributions into one of the six age groups (19-24)
    - Auto based on input table containing probability of owning a car by age compiled from available statistics
    - Equal Gender Split
    - Part time worker distribution based on age based probabilities
    - WorK industry for part time worker either service or retail based on a fixed probability
*/
Macro "Dorm Residents Synthesis"(Args)
    // Open Existing HH and Pop file after HH Resident Synthesis
    if !GetFileInfo(Args.Households) or !GetFileInfo(Args.Persons) then
        Throw("Please ensuure that the resident synthesis is complete")
    
    abm = RunMacro("Get ABM Manager", Args)
    abm.ClosePersonHHView()

    // Add new fields to HH and Person Table
    newFlds = {{Name: "UnivGQ", Type: "Short", Width: 2, Description: "1 if this is a university dorm HH (resident)"},
               {Name: "Univ", Type: "String", Width: 10, Description: "University Name (Filled from dorm synthesis)"},
               {Name: "Autos", Type: "Short", Description: "Number of autos in the HH"}}
    abm.AddHHFields(newFlds)

    newFlds = { {Name: "UnivGQStudent", Type: "Short", Width: 2, Description: "1 if this is a university student"},
                {Name: "UnivTAZ", Type: "Long", Width: 12, Description: "TAZID of university that the student belongs to"},
                {Name: "License", Type: "Short", Description: "License Status: 1 - Has driver license 2 - Does not have driver license"},
                {Name: "WorkerCategory", Type: "Short", Description: "WorkerType: 1 - FullTime 2 - PartTime"},
                {Name: "WorkDays", Type: "Short", Width: 10, Description: "Number of work days per week"},
                {Name: "WorkAttendance", Type: "Short", Width: 2, Description: "Does person work on given day? Based on value of WorkDays"},
                {Name: "AttendUniv", Type: "Short", Width: 2, Description: "Does person (non-dorm resident) attend university?|1: Yes|2: No.|Filled for Age >= 19"}}
    abm.AddPersonFields(newFlds)

    // ***** TAZ Data
    // Get relevant data from TAZ
    objT = CreateObject("Table", Args.Demographics)
    nUniv = objT.SelectByQuery({SetName: "Selection", Query: "UnivGQ = 1 and GroupQuarterPopulation > 0"})
    if nUniv = 0 then
        Throw("Check university group quarter data in TAZ Demograhics file")
    vecs = objT.GetDataVectors({FieldNames: {"TAZ", "GroupQuarterPopulation", "UnivGQ", "University", "UnivTAZ"}})
    objT = null

    // ***** Process
    // Add records
    vHHID = abm.[HH.HouseholdID]
    maxHHID = vHHID.Max()
    vPersonID = abm.[Person.PersonID]
    maxPersonID = vPersonID.Max()
    
    vecsOutHH = null
    vecsOutPersons = null
    vUnivResidents = vecs.GroupQuarterPopulation
    
    pbar = CreateObject("G30 Progress Bar", "Processing University Zones...", true, vUnivResidents.length)
    for i = 1 to vUnivResidents.length do
        nUniv = r2i(vUnivResidents[i])
        vecsDist = RunMacro("Get Univ Distributions", Args, nUniv)
        
        vHHID = Vector(nUniv, "Long", {{"Sequence", maxHHID + 1, 1}})
        vPersonID = Vector(nUniv, "Long", {{"Sequence", maxPersonID + 1, 1}})
        vOne = Vector(nUniv, "Short", {{"Constant", 1}})
        vUnivTAZ = Vector(nUniv, "Long", {{"Constant", vecs.UnivTAZ[i]}})
        vUniv = Vector(nUniv, "String", {{"Constant", vecs.University[i]}})

        // Add records, select and set HH vectors
        vecsOutHH = {TAZID: vUnivTAZ, HouseholdID: vHHID, Weight: vOne, HHSize: vOne, Autos: vecsDist.Vehicles,
                     "Inc_Category": vecsDist.IncomeCategory, "Kids_Category": vOne, "Seniors_Category": vOne,
                     UnivGQ: vOne, Univ: vUniv}
        AddRecords(abm.HHView,,,{"Empty Records": nUniv})
        abm.CreateHHSet({Filter: 'HouseholdID = null', Activate: 1})
        abm.SetHHVectors(vecsOutHH)

        // Add records, select and set Person vectors
        vecsOutPersons = {HouseholdID: vHHID, PersonID: vPersonID, Gender: vecsDist.Gender, Age: s2i(vecsDist.Age),
                          UnivGQStudent: vOne, UnivTAZ: vUnivTAZ, WorkerCategory: vecsDist.WorkerCategory,
                          WorkDays: vecsDist.WorkDays, WorkAttendance: vecsDist.WorkAttendance,
                          License: vecsDist.License, IndustryCategory: vecsDist.IndustryCategory,
                          WorkIndustry: vecsDist.WorkIndustry, AttendUniv: vOne}
        AddRecords(abm.PersonView,,,{"Empty Records": nUniv})
        abm.CreatePersonSet({Filter: 'PersonID = null', Activate: 1})
        abm.SetPersonVectors(vecsOutPersons)
        
        // Reset Max IDs
        maxHHID = vHHID[nUniv]
        maxPersonID = vPersonID[nUniv]
        if pbar.Step() then
            Return()
    end
    pbar.Destroy()
    abm.ActivateHHSet()
    abm.ActivatePersonSet()
    Return(true)
endMacro


// Macro that generates vectors for Age, Gender, IncomeCategory, WorkerStatus and Vehicle
// The number of records are passed to the macro
// The distributions are based on input probability numbers/tables
Macro "Get Univ Distributions"(Args, nUniv)
    vecsDist = null
    
    // Gender
    SetRandomSeed(42*nUniv + 1)
    fPct = Args.FemaleStudentsPct
    mPct = 100 - fPct
    params = null
    params.population = {1, 2}
    params.weight = {mPct, fPct}
    vecsDist.Gender = RandSamples(nUniv, "Discrete", params)

    // Income
    SetRandomSeed(42*nUniv + 2)
    IncCat1Pct = Args.IncCategory1Pct
    IncCat2Pct = 100 - IncCat1Pct
    params.weight = {IncCat1Pct, IncCat2Pct}
    vecsDist.IncomeCategory = RandSamples(nUniv, "Discrete", params)

    // ***** Age Related Distributions
    ageChars = Args.StudentCharacteristics
    arrAge = ageChars.Age
    arrAgePct = ageChars.Percentage
    if Sum(arrAgePct) <> 100 then
        Throw("Incorrect Age Percentages in 'StudentCharacteristics' table for dorm population synthesis. They do not sum up to 100.0")

    // Age
    params = null
    params.population = arrAge
    params.weight = arrAgePct
    vecsDist.Age = RandSamples(nUniv, "Discrete", params)

    // Vehicle Availability
    vProbVeh = Vector(nUniv, "Double",)
    vProbWrk = Vector(nUniv, "Double",)
    // Get Prob Vector First
    for i = 1 to arrAge.length do
        age = arrAge[i]
        probVeh = ageChars.[Vehicle Probability][i]
        probWrk = ageChars.[Worker Probability][i]
        vProbVeh = if vecsDist.Age = age then probVeh else vProbVeh
        vProbWrk = if vecsDist.Age = age then probWrk else vProbWrk 
    end

    SetRandomSeed(42*nUniv + 3)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.Vehicles = if vRand <= vProbVeh then 1 else 0
    vecsDist.License = if vRand <= vProbVeh then 1 else 2

    // Worker Category (PartTime or NonWorker)
    SetRandomSeed(42*nUniv + 4)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.WorkerCategory = if vRand <= vProbWrk then 2 else null
    vecsDist.WorkDays = if vecsDist.WorkerCategory = 2 then 3 else null
    
    SetRandomSeed(42*nUniv + 5)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.WorkAttendance = if vecsDist.WorkerCategory = 2 and vRand <= 0.6 then 1
                              else if vecsDist.WorkerCategory = 2 then 0
                              else null

    // Work Industry (One of Service or Retail)
    SetRandomSeed(42*nUniv + 6)
    SvcPct = Args.ServiceEmpPct
    RetPct = 100 - SvcPct
    params.weight = {SvcPct, RetPct}
    params.population = {8, 5}
    vecsDist.IndustryCategory = RandSamples(nUniv, "Discrete", params)   // Note only generating enough records for the workers
    vecsDist.IndustryCategory = if vecsDist.WorkerCategory = 2 then vecsDist.IndustryCategory else 10
    vecsDist.WorkIndustry = if vecsDist.IndustryCategory = 8 then 9 
                            else if vecsDist.IndustryCategory = 10 then 12
                            else vecsDist.IndustryCategory

    Return(vecsDist)
endMacro
