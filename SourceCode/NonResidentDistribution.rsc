/*
This macro will create trip productions and attractions for nonresident trips (externals)
*/

Macro "External DistributionAndTOD" (Args)    // Export Import

    ret_value = RunMacro("Estimate IE Trips", Args)
    if !ret_value then goto quit
    ret_value = RunMacro("Estimate EE Trips", Args)
    if !ret_value then goto quit
    quit:
    Return(ret_value)
endmacro

macro "Estimate IE Trips" (Args)
    ret_value = 1
    periods = {"AM", "MD", "PM", "NT"}
    
    dem = CreateObject("Table", Args.DemographicOutputs)
    etab = CreateObject("Table", Args.ExternalTrips)
    LineDB = Args.HighwayDatabase
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    NodeLayer = Node.GetView()

    PKSkim = Args.HighwaySkimAM
    OPSkim = Args.HighwaySkimOP
    mpk = CreateObject("Matrix", PKSkim)
    mop = CreateObject("Matrix", OPSkim)
    mskim = {mpk, mop, mpk, mop}
 
    ix_msplit = Args.IXModeSplit
    ix_split_da = ix_msplit.Split[1]
    ix_split_sr = ix_msplit.Split[2]  
    ie_vec = etab.AADT_O
    ei_vec = etab.AADT_D
    gamma_a = etab.GammaA
    gamma_b = etab.GammaB
    gamma_c = etab.GammaC
    pbar = CreateObject("G30 Progress Bar", "Creating External OD Matrices", false, 4)

    for i = 1 to periods.length do
        msk = mskim[i]
        per = periods[i]
        ExternalModel = Args.ExternalTripGen
        coeffs = ExternalModel.Coefficient
        vars = ExternalModel.Variable
        externalLR = coeffs[1] * dem.(vars[1])
        for j = 2 to vars.length do
            externalLR = externalLR + coeffs[j] * dem.(vars[j])
        end

        // Apply the fraction for the appropriate period.
        iePeriod = ie_vec * etab.("ImportedFraction" + per)
        eiPeriod = ei_vec * etab.("ExportedFraction" + per)
        iz_file = Args.(per + "ExternalTrips")
        ixobj = msk.CopyStructure({FileName: iz_file, 
                                    Cores: {"DriveAlone VehicleTrips", "Carpool VehicleTrips", "LTRK Trips", "MTRK Trips", "HTRK Trips"}, 
                                    Label: per + "ExternalTrips"})
        idx = ixobj.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
        idxInt = ixobj.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
        idxExt = ixobj.AddIndex({IndexName: "External",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'External'"})

        ixobj.SetRowIndex("External")
        ixobj.SetColIndex("InternalTAZ")
        msk.SetRowIndex("External")
        msk.SetColIndex("InternalTAZ")

        import_ids = msk.GetVector({Core: "Time", Index: "Row"})
        for j = 1 to import_ids.length do
            t = etab.PCT_AUTO_IEEI[j]
            import_vals = msk.GetVector({Core: "Time", Row: import_ids[j]})
            import_vals =  externalLR * gamma_a[j] * Pow(import_vals, gamma_b[j]) * Exp(gamma_c[j] * import_vals)
            vsum_import = VectorStatistic(import_vals, "Sum", )
            import_vals = if vsum_import = 0 then 0 else import_vals / vsum_import * iePeriod[j]
            vsum_import = VectorStatistic(import_vals, "Sum", )
            ixobj.SetVector({Core: "DriveAlone VehicleTrips", Vector: import_vals * etab.PCT_AUTO_IEEI[j]  * ix_split_da, Row: import_ids[j]})
            ixobj.SetVector({Core: "Carpool VehicleTrips", Vector: import_vals * etab.PCT_AUTO_IEEI[j] * ix_split_sr, Row: import_ids[j]})
            ixobj.SetVector({Core: "LTRK Trips", Vector: import_vals * etab.PCT_LTRK_IEEI[j], Row: import_ids[j]})
            ixobj.SetVector({Core: "MTRK Trips", Vector: import_vals * etab.PCT_MTRK_IEEI[j], Row: import_ids[j]})
            ixobj.SetVector({Core: "HTRK Trips", Vector: import_vals * etab.PCT_HTRK_IEEI[j], Row: import_ids[j]})
        end
        ixobj.SetRowIndex("InternalTAZ")
        ixobj.SetColIndex("External")
        msk.SetRowIndex("InternalTAZ")
        msk.SetColIndex("External")
        export_ids = msk.GetVector({Core: "Time", Index: "Column"})
        for j = 1 to export_ids.length do
            export_vals = msk.GetVector({Core: "Time", Column: export_ids[j]})
            export_vals.rowbased = True
            export_vals =  externalLR * gamma_a[j] * Pow(export_vals, gamma_b[j]) * Exp(gamma_c[j] * export_vals)
            vsum_export = VectorStatistic(export_vals, "Sum", )
            export_vals = if vsum_export = 0 then 0 else export_vals / vsum_export * eiPeriod[j]
            vsum_export = VectorStatistic(export_vals, "Sum", )
            export_vals.rowbased = False
            ixobj.SetVector({Core: "DriveAlone VehicleTrips", Vector: export_vals * etab.PCT_AUTO_IEEI[j] * ix_split_da, Column: export_ids[j]})
            ixobj.SetVector({Core: "Carpool VehicleTrips", Vector: export_vals * etab.PCT_AUTO_IEEI[j] * ix_split_sr, Column: export_ids[j]})
            ixobj.SetVector({Core: "LTRK Trips", Vector: export_vals * etab.PCT_LTRK_IEEI[j], Column: export_ids[j]})
            ixobj.SetVector({Core: "MTRK Trips", Vector: export_vals * etab.PCT_MTRK_IEEI[j], Column: export_ids[j]})
            ixobj.SetVector({Core: "HTRK Trips", Vector: export_vals * etab.PCT_HTRK_IEEI[j], Column: export_ids[j]})
        end
/*
        // incorporate into OD trip matrix
        idxnames = ixobj.GetIndexNames()
        ixobj.SetRowIndex(idxnames[1][1])
        ixobj.SetColIndex(idxnames[2][1])
        mOD = CreateObject("Matrix", Args.(per + "_OD"))
        mOD.[DriveAlone VehicleTrips] := nz(mOD.[DriveAlone VehicleTrips]) + nz(idxobj.[DriveAlone VehicleTrips])
        mOD.[Carpool VehicleTrips] := nz(mOD.[Carpool VehicleTrips]) + nz(idxobj.[Carpool VehicleTrips])
        mOD.[LTRK Trips] := nz(mOD.[LTRK Trips]) + nz(idxobj.[LTRK Trips])
        mOD.[MTRK Trips] := nz(mOD.[MTRK Trips]) + nz(idxobj.[MTRK Trips])
        mOD.[HTRK Trips] := nz(mOD.[HTRK Trips]) + nz(idxobj.[HTRK Trips])
        idxnames = msk.GetIndexNames()
        msk.SetRowIndex(idxnames[1][1])
        msk.SetColIndex(idxnames[2][1])
*/
        pbar.Step()

    end

   quit:    
    Return(ret_value)

endmacro



Macro "Estimate EE Trips"(Args)

    ret_value = 1
    // Inputs
    seed_matrix = Args.SeedEEMatrix
    exttrips = Args.ExternalTrips

    // Parameters
    ee_msplit = Args.EE_DA_SR_Factors
    ee_split_da = ee_msplit.Split[1]
    ee_split_sr = ee_msplit.Split[2]  

 
    // Outputs
    ee_matrix = Args.EEMatrix

    // Run the growth factor method for the daily EE Matrix
    temptab = GetTempFileName("*.bin")
    CopyTableFiles(, "FFB", exttrips, , temptab, )
    etab = CreateObject("Table", temptab)
    fields = {
    {FieldName: "AADT_O_Ext"}, 
    {FieldName: "AADT_D_Ext"}
    }
    etab.AddFields({Fields: fields})
    
    v1 = etab.AADT_O * (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    v2 = etab.AADT_D * (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    stat1 = VectorStatistic(v1, "Sum",)
    stat2 = VectorStatistic(v2, "Sum",)
    balanced_sum = (stat1 + stat2)/2
    etab.AADT_O_Ext = v1 * (balanced_sum/stat1)
    etab.AADT_D_Ext = v2 * (balanced_sum/stat2)

    // Run Growth Factor
    obj = CreateObject("Distribution.IPF")
    obj.ResetPurposes()
    obj.DataSource = {TableName:  temptab}
    obj.BaseMatrix = {MatrixFile:  seed_matrix}
    obj.AddPurpose({Name:  "DailyFlow", Production:  "AADT_O_Ext", Attraction:  "AADT_D_Ext" })
    obj.OutputMatrix({MatrixFile: ee_matrix, Matrix: "ExternalTrips"})
    ret_value = obj.Run()
    if !ret_value then goto quit
 
    // Open the EE matrix and determine the AM, PM, MD and NT matrices
    periods = {"AM", "MD", "PM", "NT"}
    modes = {"DA", "SR", "LTRK", "MTRK", "HTRK"}
    mee = CreateObject("Matrix", ee_matrix)
    cores = null
    for per in periods do
        for mode in modes do
            cores = cores + {per + " " + mode + " Flow"}
        end
        cores = cores + {per + " Flow"}
    end
    // input fraction of external flow for time period and for import and export
    mee.AddCores(cores)
    AMImp = etab.ImportedFractionAM
    AMExp = etab.ExportedFractionAM
    AMExp.rowbased = True
    PMImp = etab.ImportedFractionPM
    PMExp = etab.ExportedFractionPM
    PMExp.rowbased = True
    MDImp = etab.ImportedFractionMD
    MDExp = etab.ExportedFractionMD
    MDExp.rowbased = True
    NTImp = etab.ImportedFractionNT
    NTExp = etab.ExportedFractionNT
    NTExp.rowbased = True
    // determine flow by time period
    mee.("AM Flow") := mee.DailyFlow * AMImp + mee.DailyFlow * AMExp
    mee.("PM Flow") := mee.DailyFlow * PMImp + mee.DailyFlow * PMExp
    mee.("MD Flow") := mee.DailyFlow * MDImp + mee.DailyFlow * MDExp
    mee.("NT Flow") := mee.DailyFlow * NTImp + mee.DailyFlow * NTExp
    FractionAuto = etab.PCT_AUTO_EE / (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    FractionLTRK = etab.PCT_LTRK_EE / (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    FractionMTRK = etab.PCT_MTRK_EE / (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    FractionHTRK = etab.PCT_HTRK_EE / (etab.PCT_AUTO_EE + etab.PCT_LTRK_EE + etab.PCT_MTRK_EE + etab.PCT_HTRK_EE)
    for per in periods do
        // split into DA, SR, and Truck
        mee.(per + " DA Flow")       := mee.(per + " Flow") * FractionAuto * ee_split_da
        mee.(per + " SR Flow")       := mee.(per + " Flow") * FractionAuto * ee_split_sr
        mee.(per + " LTRK Flow")    := mee.(per + " Flow") * FractionLTRK
        mee.(per + " MTRK Flow")    := mee.(per + " Flow") * FractionMTRK
        mee.(per + " HTRK Flow")    := mee.(per + " Flow") * FractionHTRK
    /*
        // incorporate into overall trip matrices
        mOD = CreateObject("Matrix", Args.(per + "_OD"))
        mOD.SetRowIndex("External")
        mOD.SetColIndex("External")
        mOD.[DriveAlone VehicleTrips] := nz(mOD.[DriveAlone VehicleTrips]) + nz(mee.(per + " DA Flow"))
        mOD.[Carpool VehicleTrips] := nz(mOD.[Carpool VehicleTrips]) + nz(mee.(per + " SR Flow"))
        mOD.[LTRK Trips] := nz(mOD.[LTRK Trips]) + nz(mee.(per + " LTRK Flow"))
        mOD.[MTRK Trips] := nz(mOD.[MTRK Trips]) + nz(mee.(per + " MTRK Flow"))
        mOD.[HTRK Trips] := nz(mOD.[HTRK Trips]) + nz(mee.(per + " HTRK Flow"))
    */
    end

    
  quit:
    Return(ret_value)

endMacro

Macro "OtherModel"(Args)
    ret_value = 1
    quit:
    Return(ret_value)
endmacro