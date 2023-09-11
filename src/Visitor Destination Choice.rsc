/*

*/

Macro "Visitor Calculate DC" (Args)

    if Args.Iteration = 1 then do
        RunMacro("DC Size Terms", Args)
        RunMacro("Create Visitor Clusters", Args)
    end
    // HB models
    // RunMacro("Visitor DC", Args)
    // RunMacro("Apportion Visitor Trips", Args)
    // NHB model
    RunMacro("Visitor Scale NHB Size", Args)
    Throw()
    RunMacro("Visitor DC", Args, nhb = "true")
    RunMacro("Apportion Visitor Trips", Args, nhb = "true")
    return(1)
endmacro

Macro "Apply Visitor MC/DC Probabilities" (Args)
    RunMacro("Apportion Resident HB Trips", Args)
    return(1)
endmacro

/*
Creates sum product fields using DC size coefficients. Then takes the log
of those fields so it can be fed directly into the DC utility equation.

Note: this generates initial NHB size variables, but these are scaled
based on HB trip ends after those DC models are run.
*/

Macro "DC Size Terms" (Args)
    se_file = Args.DemographicOutputs
    input_dir = Args.[Input Folder]
    coeff_file = input_dir + "/visitors/dc/dc_size_terms.csv"

    sizeSpec = {DataFile: se_file, CoeffFile: coeff_file}
    RunMacro("Compute Size Terms", sizeSpec)
endmacro


// Generic size term computation macro, given the view with all relevnt fields and the coeff file. Add fields to the view.
Macro "Compute Size Terms"(sizeSpec)
    coeff_file = sizeSpec.CoeffFile
    se_vw = OpenTable("Data", "FFB", {sizeSpec.DataFile})

    // Calculate the size term fields using the coefficient file
    {drive, folder, name, ext} = SplitPath(coeff_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: coeff_file,
        field_desc: "Visitor DC Size Terms|" +
        "These are already log transformed and used directly by the DC model.|" +
        "See " + name + ext + " for details."
    })

    // Log transform the results and set any 0s to nulls
    coeff_vw = OpenTable("coeff", "CSV", {coeff_file})
    {field_names, } = GetFields(coeff_vw, "All")
    CloseView(coeff_vw)
    
    // Remove the first and last fields ("Field" and "Description")
    field_names = ExcludeArrayElements(field_names, 1, 1)
    field_names = ExcludeArrayElements(field_names, field_names.length, 1)
    input = GetDataVectors(se_vw + "|", field_names, {OptArray: TRUE})
    for field_name in field_names do
        output.(field_name) = if input.(field_name) = 0
            then null
            else Log(1 + input.(field_name))
    end
    SetDataVectors(se_vw + "|", output, )
    CloseView(se_vw)
endMacro

/*
The visitor model uses the Cluster field from the se data, but
collapses district 3 into 1 due to a lack of observations in the
survey.
*/

Macro "Create Visitor Clusters" (Args)
    se_file = Args.DemographicOutputs
    se = CreateObject("Table", se_file)
    se.AddField({FieldName: "VisCluster", type: "integer"})
    se.AddField({FieldName: "VisClusterName", type: "string"})
    se.VisCluster = if se.Cluster = 3
        then 1
        else se.Cluster
    se.VisClusterName = "c" + String(se.VisCluster)
endmacro

/*

*/

Macro "Visitor DC" (Args, nhb)
    if nhb
        then trip_types = {"NHB"}
        else trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW"}
    RunMacro("Calculate Destination Choice", Args, trip_types)
endmacro

/*

*/

Macro "Visitor Scale NHB Size" (Args)
    
    se_file = Args.DemographicOutputs
    periods = Args.TimePeriods
    trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW"}
    out_dir = Args.[Output Folder]
    trip_dir = out_dir + "/visitors/trip_matrices"

    // Add up all HB trip ends
    for trip_type in trip_types do
        for i = 1 to periods.length do
            period = periods[i][1]

            mtx_file = trip_dir + "/pa_per_trips_" + trip_type + "_" + period + ".mtx"
            mtx = CreateObject("Matrix", mtx_file)
            v_b = mtx.GetVector({Core: "dc_business", Marginal: "Column Sum"})
            if TypeOf(v_trip_ends) = "null" then do
                v_trip_ends = Vector(v_b.length, "double", {Constant: 0})
            end
            v_trip_ends = v_trip_ends + nz(v_b)
            if trip_type <> "HBW" then do
                v_p = mtx.GetVector({Core: "dc_personal", Marginal: "Column Sum"})
                v_trip_ends = v_trip_ends + nz(v_p)
            end
        end
    end
    
    // Add to SE table
    se = CreateObject("Table", se_file)
    se.AddField({
        FieldName: "vis_hb_tripends", 
        Description: "Visitor trip ends from HB purposes.|Used to scale vis NHB trips."
    })
    se.AddField({
        FieldName: "vis_hb_frac", 
        Description: "Visitor hb trip ends as a fraction of column total"
    })
    se.vis_hb_tripends = v_trip_ends
    se.vis_hb_frac = v_trip_ends / v_trip_ends.sum()
    
    // Create scaled size fields
    total_size = se.vis_NHB_Size.sum()
    v_temp = se.vis_NHB_Size * se.vis_hb_frac
    v_scaled_size = v_temp * (total_size / v_temp.sum())
    se.AddField("vis_NHB_size_scaled")
    se.vis_NHB_size_scaled = v_scaled_size    
endmacro

/*

*/

Macro "Calculate Destination Choice" (Args, trip_types)

    scen_dir = Args.[Scenario Folder]
    skims_dir = scen_dir + "\\output\\skims\\"
    input_dir = Args.[Input Folder]
    input_dc_dir = input_dir + "/visitors/dc"
    output_dir = Args.[Output Folder] + "/visitors/dc"
    se_file = Args.DemographicOutputs
    periods = Args.TimePeriods

    opts = null
    opts.output_dir = output_dir
    opts.primary_spec = {Name: "sov_skim"}
    for trip_type in trip_types do
        if Lower(trip_type) = "hbw"
            then segments = {"business"}
            else segments = {"business", "personal"}
        opts.trip_type = trip_type
        opts.zone_utils = input_dc_dir + "/" + trip_type + "_zone.csv"
        opts.cluster_data = input_dc_dir + "/" + trip_type + "_cluster.csv"
        
        for i = 1 to periods.length do
            period = periods[i][1]
            opts.period = period
            
            // Determine which sov skim to use
            sov_skim = skims_dir + "\\HighwaySkim" + period + ".mtx"
            
            // Set sources
            opts.tables = {
                se: {File: se_file, IDField: "TAZ"}
            }
            opts.cluster_equiv_spec = {File: se_file, ZoneIDField: "TAZ", ClusterIDField: "VisCluster"}
            opts.dc_spec = {DestinationsSource: "sov_skim", DestinationsIndex: "Destination"}
            for segment in segments do
                opts.segments = {segment}
                opts.matrices = {
                    intra_zonal: {File: skims_dir + "/IntraZonal.mtx"},
                    sov_skim: {File: sov_skim}
                }
                
                // RunMacro("Parallel.SetMaxEngines", 3)
                // task = CreateObject("Parallel.Task", "DC Runner", GetInterface())
                // task.Run(opts)
                // tasks = tasks + {task}
                
                // To run this code in series (and not in parallel), comment out the "task"
                // and "monitor" lines of code. Uncomment the two lines below. This can be
                // helpful for debugging.
                obj = CreateObject("NestedDC", opts)
                obj.Run()
            end
        end
    end

    // monitor = CreateObject("Parallel.TaskMonitor", tasks)
    // monitor.DisplayStatus()
    // monitor.WaitForAll()
    // if monitor.IsFailed then Throw("MC Failed")
    // monitor.CloseStatusDbox()
endmacro

Macro "DC Runner" (opts)
    obj = CreateObject("NestedDC", opts)
    obj.Run()
endmacro

/*
Note: could re-factor this along with the "Apportion Resident HB Trips" macro
to avoid duplicate code.
*/

// Macro "Update Shadow Price" (Args)
    
//     se_file = Args.SE
//     out_dir = Args.[Output Folder]
//     dc_dir = out_dir + "/resident/dc"
//     sp_file = Args.ShadowPrices
//     periods = Args.TimePeriods

//     se_vw = OpenTable("se", "FFB", {se_file})
//     sp_vw = OpenTable("sp", "FFB", {sp_file})

//     trip_type = "W_HB_W_All"
//     segments = {"v0", "ilvi", "ilvs", "ihvi", "ihvs"}

//     v_sp = GetDataVector(sp_vw + "|", "hbw", )
//     v_attrs = GetDataVector(se_vw + "|", "w_hbw_a", )

//     for period in periods do
//         for segment in segments do
//             name = trip_type + "_" + segment + "_" + period
//             dc_mtx_file = dc_dir + "/probabilities/probability_" + name + "_zone.mtx"
//             out_mtx_file = Substitute(dc_mtx_file, ".mtx", "_temp.mtx", )
//             if GetFileInfo(out_mtx_file) <> null then DeleteFile(out_mtx_file)
//             CopyFile(dc_mtx_file, out_mtx_file)
            
//             out_mtx = CreateObject("Matrix", out_mtx_file)
//             cores = out_mtx.GetCores()
            
//             v_prods = nz(GetDataVector(se_vw + "|", name, ))
//             v_prods.rowbased = "false"
//             cores.final_prob := cores.final_prob * v_prods
//             v_trips = out_mtx.GetVector({"Core": "final_prob", Marginal: "Column Sum"})
//             v_total_trips = nz(v_total_trips) + v_trips

//             out_mtx = null
//             cores = null
//             DeleteFile(out_mtx_file)
//         end
//     end
    
//     // Calculate constant adjustmente. avoid Log(0) = -inf
//     delta = if v_attrs = 0 or v_total_trips = 0
//         then 0
//         else nz(Log(v_attrs/v_total_trips)) * .85
//     v_sp_new = v_sp + delta
//     SetDataVector(sp_vw + "|", "hbw", v_sp_new, )

//     CloseView(se_vw)
//     CloseView(sp_vw)

//     // return the %RMSE
//     o = CreateObject("Model.Statistics")
//     stats = o.rmse({Method: "vectors", Predicted: v_sp_new, Observed: v_sp})
//     prmse = stats.RelRMSE
//     return(prmse)
// endmacro

/*
With DC and MC probabilities calculated, resident trip productions can be 
distributed into zones and modes.

The 'nhb' input is used because this macro is called twice. First for HB trips
only. The results of this first call are used to scale NHB size terms. After
NHB dc runs, this is called again to apportion those trips.
*/

Macro "Apportion Visitor Trips" (Args, nhb)

    se_file = Args.DemographicOutputs
    out_dir = Args.[Output Folder]
    dc_dir = out_dir + "/visitors/dc"
    mc_dir = out_dir + "/visitors/mc"
    trip_dir = out_dir + "/visitors/trip_matrices"
    periods = Args.TimePeriods
    access_modes = Args.access_modes

    se = CreateObject("Table", se_file)

    // Create output folders
    RunMacro("Create Directory", dc_dir)
    RunMacro("Create Directory", mc_dir)
    RunMacro("Create Directory", trip_dir)

    if nhb
        then trip_types = {"NHB"}
        else trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW"}

    for i = 1 to periods.length do
        period = periods[i][1]

        for trip_type in trip_types do
            if trip_type = "HBW"
                then segments = {"business"}
                else segments = {"business", "personal"}
            
            out_mtx_file = trip_dir + "/pa_per_trips_" + trip_type + "_" + period + ".mtx"
            if GetFileInfo(out_mtx_file) <> null then DeleteFile(out_mtx_file)

            for segment in segments do
                name = trip_type + "_" + segment + "_" + period
                prod_field = "prod_v" + Left(segment, 1) + trip_type + "_" + period // e.g. prod_vbHBRec_AM
                
                dc_mtx_file = dc_dir + "/probabilities/probability_" + name + "_zone.mtx"
                dc_mtx = CreateObject("Matrix", dc_mtx_file)
                dc_cores = dc_mtx.GetCores()
                mc_mtx_file = mc_dir + "/probabilities/probability_" + name + ".mtx"
                if segment = segments[1] then do
                    CopyFile(mc_mtx_file, out_mtx_file)
                    out_mtx = CreateObject("Matrix", out_mtx_file)
                    core_names = out_mtx.GetCoreNames()
                    cores = out_mtx.GetCores()
                    for core_name in core_names do
                        cores.(core_name) := nz(cores.(core_name)) * 0
                    end
                end
                mc_mtx = CreateObject("Matrix", mc_mtx_file)
                mc_cores = mc_mtx.GetCores()

                v_prods = nz(se.(prod_field))
                v_prods.rowbased = "false"

                mode_names = mc_mtx.GetCoreNames()
                out_cores = out_mtx.GetCores()
                for mode in mode_names do
                    out_cores.(mode) := nz(out_cores.(mode)) + v_prods * nz(dc_cores.final_prob) * nz(mc_cores.(mode))
                    
                    if mode = mode_names[1] then do
                        // Add extra cores to hold dc-only results
                        dc_core = "dc_" + segment
                        out_mtx.AddCores({dc_core})
                        dc_core = out_mtx.GetCore(dc_core)
                        dc_core := nz(dc_core) + v_prods * nz(dc_cores.final_prob)
                    end
                end
            end
        end
    end
endmacro