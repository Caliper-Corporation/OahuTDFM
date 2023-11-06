/*
Oahu visitor model. These steps generate the outputs found in the
output/visitors directory. Before assignment, visitor trips by mode
are added into the main assignment OD matrices.
*/

Macro "Visitor Model" (Args)

    RunMacro("Visitor Lodging Locations", Args)
    RunMacro("Visitor Trip Generation", Args)
    RunMacro("Visitor Time of Day", Args)
    RunMacro("Visitor Create MC Features", Args)
    RunMacro("Visitor Calculate MC", Args)
    RunMacro("Visitor Calculate DC", Args)
    RunMacro("Visitor Directionality", Args)
    RunMacro("Visitor Occupancy", Args)
    return(1)
endmacro

/*
Determine visitor home/loding locations
*/

Macro "Visitor Lodging Locations" (Args)

    se_file = Args.DemographicOutputs

    se = CreateObject("Table", se_file)
    se.AddFields({Fields: {
        {FieldName: "visitors_p", Description: "Visitors in the zone on a personal trip to the island"},
        {FieldName: "visitors_b", Description: "Visitors in the zone on a business trip to the island"}
    }})

    se.visitors_b = se.OccupiedHH * Args.[Vis HH Occ Rate] * Args.[Vis HH Business Ratio] +
        se.HR * Args.[Vis Hotel Occ Rate] * Args.[Vis Hotel Business Ratio] + 
        se.RC * Args.[Vis Condo Occ Rate] * Args.[Vis Condo Business Ratio]
    se.visitors_b = se.visitors_b * Args.[Vis Business Party Size] * Args.[Vis Party Calibration Factor]
    se.visitors_p = se.OccupiedHH * Args.[Vis HH Occ Rate] * (1 - Args.[Vis HH Business Ratio]) +
        se.HR * Args.[Vis Hotel Occ Rate] * (1 - Args.[Vis Hotel Business Ratio]) + 
        se.RC * Args.[Vis Condo Occ Rate] * (1 - Args.[Vis Condo Business Ratio])
    se.visitors_p = se.visitors_p * Args.[Vis Personal Party Size] * Args.[Vis Party Calibration Factor]
endmacro

/*
Generate trips by purpose for business and personal visitors
*/

Macro "Visitor Trip Generation" (Args)
    se_file = Args.DemographicOutputs
    rate_file = Args.[Vis Trip Rates]

    se = CreateObject("Table", se_file)
    se_vw = se.GetView()
    {drive, folder, name, ext} = SplitPath(rate_file)
    RunMacro("Create Sum Product Fields", {
        view: se_vw, factor_file: rate_file,
        field_desc: "CV Productions and Attractions|See " + name + ext + " for details."
    })
endmacro

/*

*/

Macro "Visitor Time of Day" (Args)

    se_file = Args.DemographicOutputs
    input_dir = Args.[Input Folder]
    tod_file = input_dir + "\\visitors\\vis_tod_factors.csv"

    se_vw = OpenTable("per", "FFB", {se_file})
    fac_vw = OpenTable("tod_fac", "CSV", {tod_file})
    v_purp = GetDataVector(fac_vw + "|", "trip_purp", )
    v_tod = GetDataVector(fac_vw + "|", "tod", )
    v_fac = GetDataVector(fac_vw + "|", "factor", )

    for i = 1 to v_purp.length do
        purp = v_purp[i]
        tod = v_tod[i]
        fac = v_fac[i]

        if purp = "HBW" 
            then segments = {"business"}
            else segments = {"business", "personal"}

        for segment in segments do
            daily_name = "prod_v" + Left(segment, 1) + purp // e.g. prod_vbHBRec
            v_daily = GetDataVector(se_vw + "|", daily_name, )
            v_result = v_daily * fac
            field_name = daily_name + "_" + tod
            a_fields_to_add = a_fields_to_add + {
                {field_name, "Real", 10, 2,,,, "Visitor productions by TOD"}
            }
            data.(field_name) = v_result
        end
    end
    RunMacro("Add Fields", {view: se_vw, a_fields: a_fields_to_add})
    SetDataVectors(se_vw + "|", data, )    
    CloseView(se_vw)
    CloseView(fac_vw)
endmacro

/*
Create any features needed by the visitor mc model
*/

Macro "Visitor Create MC Features" (Args)
    se_file = Args.DemographicOutputs

    se = CreateObject("Table", se_file)
    se.AddField("AreaTypeNum")
    se.AreaTypeNum = if se.AreaType = "Downtown" then 1
        else if se.AreaType = "Urban" then 2
        else if se.AreaType = "Suburban" then 3
        else if se.AreaType = "Rural" then 4
endmacro

/*
Loops over purposes and preps options for the "MC" macro
*/

Macro "Visitor Calculate MC" (Args)

    scen_dir = Args.[Scenario Folder]
    skims_dir = scen_dir + "\\output\\skims\\"
    input_dir = Args.[Input Folder]
    input_mc_dir = input_dir + "/visitors/mc"
    output_dir = Args.[Output Folder] + "/visitors/mc"
    periods = {"AM", "PM", "OP"}
    mode_table = Args.TransitModeTable
    access_modes = Args.AccessModes
    se_file = Args.DemographicOutputs
    
    // Specify trip purposes and transit modes
    trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW", "NHB"}
    transit_modes = RunMacro("Get Transit Net Def Col Names", mode_table)
    pos = transit_modes.position("all")
    if pos > 0 then transit_modes = ExcludeArrayElements(transit_modes, pos, 1)

    opts = null
    opts.primary_spec = {Name: "bus_skim"}
    for trip_type in trip_types do
        opts.segments = {"personal", "business"}
        opts.trip_type = trip_type
        opts.util_file = input_mc_dir + "/" + trip_type + ".csv"
        nest_file = input_mc_dir + "/" + trip_type + "_nest.csv"
        if GetFileInfo(nest_file) <> null 
            then opts.nest_file = nest_file
            else opts.nest_file = null

        // Set sources
        opts.tables = {
            se: {File: se_file, IDField: "TAZ"}
        }

        for period in periods do
            opts.period = period
            opts.matrices = {
                auto_skim: {File: skims_dir + "\\HighwaySkim" + period + ".mtx"},
                walk_skim: {File: skims_dir + "\\WalkSkim.mtx"},
                bike_skim: {File: skims_dir + "\\BikeSkim.mtx"},
                iz_skim: {File: skims_dir + "\\IntraZonal.mtx"}
            }
            // Transit skims depend on which modes are present in the scenario
            for transit_mode in transit_modes do
                source_name = transit_mode + "_skim"
                file_name = skims_dir + "transit\\" + period + "_w_" + transit_mode + ".mtx"
                if GetFileInfo(file_name) <> null then opts.matrices.(source_name) = {File: file_name}
            end

            opts.output_dir = output_dir
            RunMacro("MC", opts)
        end
    end
endmacro

/*
Calculates the visitor DC models and applies probabilities to get trip tables.
Does HB first and NHB second.
*/

Macro "Visitor Calculate DC" (Args)

    if Args.Iteration = 1 then do
        RunMacro("DC Size Terms", Args)
        RunMacro("Create Visitor Clusters", Args)
    end
    // HB models
    RunMacro("Visitor DC", Args)
    RunMacro("Visitor Apply Probabilities", Args)
    // NHB model
    nhb = "true"
    RunMacro("Visitor Scale NHB Productions", Args)
    RunMacro("Visitor DC", Args, nhb)
    RunMacro("Visitor Apply Probabilities", Args, nhb)
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
The total visitor NHB productions are kept from trip gen, but the locations
are scaled to match HB trip attractions from HB destination choice. This
ensures that visitor NHB trips happen near visitor HB trips.
*/

Macro "Visitor Scale NHB Productions" (Args)
    
    se_file = Args.DemographicOutputs
    periods = {"AM", "PM", "OP"}
    trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW"}
    out_dir = Args.[Output Folder]
    trip_dir = out_dir + "/visitors/trip_matrices"

    se = CreateObject("Table", se_file)

    // For each period, add up all HB attractions and set them to NHB prods
    // after scaling.
    for period in periods do
        for segment in {"business", "personal"} do
            v_hb_attrs = null
            
            // Add up HB attractions    
            for trip_type in trip_types do
                if trip_type = "HBW" and segment = "personal" then continue

                mtx_file = trip_dir + "/pa_per_trips_" + trip_type + "_" + period + ".mtx"
                mtx = CreateObject("Matrix", mtx_file)

                v = mtx.GetVector({Core: "dc_" + segment, Marginal: "Column Sum"})
                if TypeOf(v_hb_attrs) = "null" then do
                    v_hb_attrs = Vector(v.length, "double", {Constant: 0})
                end
                v_hb_attrs = v_hb_attrs + nz(v)
            end

            // Scale HB attractions to match total NHB productions and set
            // them as NHB productions.
            // NHB prod fields look like "prod_vbNHB_OP"
            bp = Left(segment, 1)
            field_name = "prod_v" + bp + "NHB_" + period
            factor = se.(field_name).sum() / v_hb_attrs.sum()
            se.(field_name) = v_hb_attrs * factor
        end
    end
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
    periods = {"AM", "PM", "OP"}

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
        
        for period in periods do
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

Macro "Visitor Apply Probabilities" (Args, nhb)

    se_file = Args.DemographicOutputs
    out_dir = Args.[Output Folder]
    dc_dir = out_dir + "/visitors/dc"
    mc_dir = out_dir + "/visitors/mc"
    trip_dir = out_dir + "/visitors/trip_matrices"
    periods = {"AM", "PM", "OP"}
    access_modes = Args.access_modes

    se = CreateObject("Table", se_file)

    // Create output folders
    RunMacro("Create Directory", dc_dir)
    RunMacro("Create Directory", mc_dir)
    RunMacro("Create Directory", trip_dir)

    if nhb
        then trip_types = {"NHB"}
        else trip_types = {"HBEat", "HBO", "HBRec", "HBShop", "HBW"}

    for period in periods do
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

/*
Convert from PA to OD format
*/

Macro "Visitor Directionality" (Args)

    out_dir = Args.[Output Folder]
    trip_dir = out_dir + "/visitors/trip_matrices"
    factor_file = Args.VisDirectionFactors

    factors = CreateObject("Table", factor_file)
    fac_vw = factors.GetView()
    rh = GetFirstRecord(fac_vw + "|", )
    modes = {"auto", "tnc", "bus"}
    while rh <> null do
        trip_type = fac_vw.trip_type
        period = fac_vw.tod
        pa_factor = fac_vw.pa_fac

        pa_mtx_file = trip_dir + "/pa_per_trips_" + trip_type + "_" + period + ".mtx"
        // they are actually still person trips at the end of this macro,
        // but the occupancy macro follows directly after and converts them to
        // vehicle trips.
        od_mtx_file = trip_dir + "/od_veh_trips_" + trip_type + "_" + period + ".mtx"
        CopyFile(pa_mtx_file, od_mtx_file)

        mtx = CreateObject("Matrix", od_mtx_file)
        t_mtx = mtx.Transpose()

        core_names = mtx.GetCoreNames()
        for mode in modes do
            if core_names.position(mode) = 0 then continue
            mtx.(mode) := mtx.(mode) * pa_factor + t_mtx.(mode) * (1 - pa_factor)
        end

        // Drop non-auto modes (these remain PA format)
        core_names = mtx.GetCoreNames()
        for core_name in core_names do
            if modes.position(core_name) = 0 then mtx.DropCores({core_name})
        end
        
        rh = GetNextRecord(fac_vw + "|", rh, )
    end
endmacro

/*
Split auto core into sov/hov
*/

Macro "Visitor Occupancy" (Args)

    out_dir = Args.[Output Folder]
    trip_dir = out_dir + "/visitors/trip_matrices"
    factor_file = Args.VisOccupancyFactors

    factors = CreateObject("Table", factor_file)
    fac_vw = factors.GetView()
    rh = GetFirstRecord(fac_vw + "|", )
    auto_modes = {"auto"}
    while rh <> null do
        trip_type = fac_vw.trip_type
        period = fac_vw.tod
        pct_sov = fac_vw.pct_sov
        hov_occ = fac_vw.hov_occ

        od_mtx_file = trip_dir + "/od_veh_trips_" + trip_type + "_" + period + ".mtx"
        mtx = CreateObject("Matrix", od_mtx_file)
        mtx.AddCores({"sov", "hov"})
        mtx.sov := mtx.auto * pct_sov
        mtx.hov := (mtx.auto - mtx.sov) / hov_occ
        
        rh = GetNextRecord(fac_vw + "|", rh, )
    end
endmacro