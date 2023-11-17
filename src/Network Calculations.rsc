/*

*/

Macro "Network Calculations" (Args)
    RunMacro("CopyDataToOutputFolder", Args)
    RunMacro("Filter Transit Modes", Args)
    RunMacro("Expand DTWB", Args)
    RunMacro("Mark KNR Nodes", Args)
    RunMacro("Determine Area Type", Args)
    RunMacro("Add Cluster Info to SE Data", Args)
    RunMacro("Create Visitor Clusters", Args)
    RunMacro("Speeds and Capacities", Args)
    RunMacro("CalculateTransitSpeeds Oahu", Args)
    RunMacro("Compute Intrazonal Matrix", Args)
    RunMacro("Create Intra Cluster Matrix", Args)

    return(1)
endmacro

/*
Remove network columns from the mode table that if that mode doesn't exist in 
the scenario. This will in turn control which networks (tnw) get created.
*/

Macro "Filter Transit Modes" (Args)
    rts_file = Args.TransitRoutes
    mode_file = Args.TransitModeTable

    // Open a map and determine if rail is present
    map = CreateObject("Map", rts_file)
    {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
    n = map.SelectByQuery({
        SetName: "rail",
        Query: "Mode = 7"
    })
    map = null

    if n = 0 then do
        // Can't edit CSVs directly, so convert to a MEM table
        temp = CreateObject("Table", mode_file)
        tbl = temp.Export({ViewName: "temp"})
        tbl.DropFields("rail")
        temp = null
        tbl.Export({FileName: mode_file})
    end

    // Remove modes from MC parameter files
    RunMacro("Filter Visitor Transit Modes", Args)
endmacro

/*
Removes modes from the visitor csv parameter files if they don't exist
in the scenario.
*/

Macro "Filter Visitor Transit Modes" (Args)
    
    mode_table = Args.TransitModeTable
    // access_modes = Args.access_modes
    mc_dir = Args.[Input Folder] + "/visitors/mc"
    
    transit_modes = RunMacro("Get Transit Net Def Col Names", mode_table)

    // get visitor trip purposes
    factor_file = Args.VisOccupancyFactors
    fac_tbl = CreateObject("Table", factor_file)
    trip_types = fac_tbl.trip_type
    trip_types = SortVector(trip_types, {Unique: "true"})

    for trip_type in trip_types do
        coef_file = mc_dir + "/" + trip_type + ".csv"
        coef_tbl = CreateObject("Table", coef_file)

        // Start by selecting all non-transit modes
        coef_tbl.SelectByQuery({
            SetName: "export",
            Query: "Alternative = 'auto' or Alternative = 'tnc' or Alternative = 'walk' or Alternative = 'bike'"
        })
        // Now add transit modes that exist to the selection
        for mode in transit_modes do
            coef_tbl.SelectByQuery({
                SetName: "export",
                Operation: "more",
                Query: "Alternative = '" + mode + "'"
            })  
        end
        temp = coef_tbl.Export({ViewName: "temp"})
        coef_tbl = null
        temp.Export({FileName: coef_file})
    end
endmacro

/*
The DTWB field has a string that says whether

Drive
Transit
Walk
Bike

is available. This macro converts that into one-hot fields.
*/

Macro "Expand DTWB" (Args)
    
    hwy_dbd = Args.HighwayDatabase

    tbl = CreateObject("Table", {FileName: hwy_dbd, Layer: 2})
    tbl.AddField("D")
    tbl.AddField("T")
    tbl.AddField("W")
    tbl.AddField("B")
    
    v_dtwb = tbl.DTWB
    set = null
    set.D = if Position(v_dtwb, "D") <> 0 then 1 else 0
    set.T = if Position(v_dtwb, "T") <> 0 then 1 else 0
    set.W = if Position(v_dtwb, "W") <> 0 then 1 else 0
    set.B = if Position(v_dtwb, "B") <> 0 then 1 else 0
    tbl.SetDataVectors({FieldData: set})
endmacro

/*

*/

Macro "Mark KNR Nodes" (Args)

    rts_file = Args.TransitRoutes

    map = CreateObject("Map", rts_file)
    {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()

    node = CreateObject("Table", nlyr)
    node.AddField({
        FieldName: "KNR",
        Description: "If node can be considered for KNR|(If a bus stop is at the node)"
    })
    links = CreateObject("Table", llyr)
    links.SelectByQuery({
        SetName: "drive_links",
        Query: "Select * where D  = 1"
    })
    SetLayer(nlyr)
    SelectByLinks("drive nodes", "several", "drive_links", )
    n = SelectByVicinity ("knr", "several", slyr + "|", 10/5280, {"Source And": "drive nodes"})
    v = Vector(n, "Long", {Constant: 1})
    node.ChangeSet("knr")
    node.KNR = v

endmacro

/*
Prepares input options for the AreaType.rsc library of tools, which
tags TAZs and Links with area types.
*/

Macro "Determine Area Type" (Args)

    scen_dir = Args.[Scenario Folder]
    taz_dbd = Args.TAZGeography
    se_bin = Args.DemographicOutputs
    hwy_dbd = Args.HighwayDatabase
    area_tbl = Args.AreaTypes

    // Get area from TAZ layer
    {map, {taz_lyr}} = RunMacro("Create Map", {file: taz_dbd})

    // Calculate total employment and density
    se_vw = OpenTable("se", "FFB", {se_bin, })
    a_fields =  {
        {"TotalEmployment", "Integer", 10, ,,,, "Total employment"},
        {"Density", "Real", 10, 2,,,, "Density used in area type calculation.|Considers HH and Emp."},
        {"AreaType", "Character", 10,,,,, "Area Type"},
        {"ATSmoothed", "Integer", 10,,,,, "Whether or not the area type was smoothed"},
        {"EmpDensity", "Real", 10, 2,,,, "Employment density. Used in some DC models.|TotalEmp / Area."}
    }
    RunMacro("Add Fields", {view: se_vw, a_fields: a_fields})

    // Join the se to TAZ
    jv = JoinViews("jv", taz_lyr + ".TAZID", se_vw + ".TAZ", )

    data = GetDataVectors(
        jv + "|",
        {
            "Area",
            "Population",
            "GroupQuarterPopulation",
            "Emp_Agriculture",
            "Emp_Manufacturing",
            "Emp_Wholesale",
            "Emp_Retail",
            "Emp_TransportConstruction",
            "Emp_FinanceRealEstate",
            "Emp_Education",
            "Emp_HealthCare",
            "Emp_Services",
            "Emp_Public",
            "Emp_Hotel",
            "Emp_Military"
        },
        {OptArray: TRUE, "Missing as Zero": TRUE}
    )
    tot_emp = data.Emp_Agriculture + 
        data.Emp_Manufacturing + 
        data.Emp_Wholesale + 
        data.Emp_Retail + 
        data.Emp_TransportConstruction + 
        data.Emp_FinanceRealEstate + 
        data.Emp_Education + 
        data.Emp_HealthCare + 
        data.Emp_Services + 
        data.Emp_Public + 
        data.Emp_Hotel + 
        data.Emp_Military
    data.HH_POP = data.Population - data.GroupQuarterPopulation
    factor = data.HH_POP.sum() / tot_emp.sum()
    density = (data.HH_POP + tot_emp * factor) / data.area
    emp_density = tot_emp / data.area
    areatype = Vector(density.length, "String", )
    for i = 1 to area_tbl.length do
        name = area_tbl[i].AreaType
        cutoff = area_tbl[i].Density
        areatype = if density >= cutoff then name else areatype
    end
    // SetDataVector(jv + "|", "TotalEmp", tot_emp, )
    SetDataVector(jv + "|", se_vw + ".TotalEmployment", tot_emp, )
    SetDataVector(jv + "|", se_vw + ".Density", density, )
    SetDataVector(jv + "|", se_vw + ".AreaType", areatype, )
    SetDataVector(jv + "|", se_vw + ".EmpDensity", emp_density, )

    views.se_vw = se_vw
    views.jv = jv
    views.taz_lyr = taz_lyr
    RunMacro("Smooth Area Type", Args, map, views)
    RunMacro("Tag Highway with Area Type", Args, map, views)

    CloseView(jv)
    CloseView(se_vw)
    CloseMap(map)
EndMacro

/*
Uses buffers to smooth the boundaries between the different area types.
*/

Macro "Smooth Area Type" (Args, map, views)
    
    taz_dbd = Args.TAZGeography
    area_tbl = Args.AreaTypes
    se_vw = views.se_vw
    jv = views.jv
    taz_lyr = views.taz_lyr

    // This smoothing operation uses Enclosed inclusion
    if GetSelectInclusion() = "Intersecting" then do
        reset_inclusion = TRUE
        SetSelectInclusion("Enclosed")
    end

    // Loop over the area types in reverse order (e.g. Urban to Rural)
    // Skip the last (least dense) area type (usually "Rural") as those do
    // not require buffering.
    for t = area_tbl.length to 2 step -1 do
        type = area_tbl[t].AreaType
        buffer = area_tbl[t].Buffer

        // Select TAZs of current type
        SetView(jv)
        query = "Select * where " + se_vw + ".AreaType = '" + type + "'"
        n = SelectByQuery("selection", "Several", query)

        if n > 0 then do
            // Create a temporary buffer (deleted at end of macro)
            // and add to map.
            a_path = SplitPath(taz_dbd)
            bufferDBD = a_path[1] + a_path[2] + "ATbuffer.dbd"
            CreateBuffers(bufferDBD, "buffer", {"selection"}, "Value", {buffer},)
            bLyr = AddLayer(map,"buffer",bufferDBD,"buffer")

            // Select zones within the 1 mile buffer that have not already
            // been smoothed.
            SetLayer(taz_lyr)
            n2 = SelectByVicinity("in_buffer", "several", "buffer|", , )
            qry = "Select * where ATSmoothed = 1"
            n2 = SelectByQuery("in_buffer", "Less", qry)

            if n2 > 0 then do
            // Set those zones' area type to the current type and mark
            // them as smoothed
            opts = null
            opts.Constant = type
            v_atype = Vector(n2, "String", opts)
            opts = null
            opts.Constant = 1
            v_smoothed = Vector(n2, "Long", opts)
            SetDataVector(
                jv + "|in_buffer", se_vw + "." + "AreaType", v_atype,
            )
            SetDataVector(
                jv + "|in_buffer", se_vw + "." + "ATSmoothed", v_smoothed,
            )
            end

            DropLayer(map, bLyr)
            DeleteDatabase(bufferDBD)
        end
    end

    if reset_inclusion then SetSelectInclusion("Intersecting")
EndMacro

/*
Tags highway links with the area type of the TAZ they are nearest to.
*/

Macro "Tag Highway with Area Type" (Args, map, views)

    hwy_dbd = Args.HighwayDatabase
    area_tbl = Args.AreaTypes
    se_vw = views.se_vw
    jv = views.jv
    taz_lyr = views.taz_lyr

    // This smoothing operation uses intersecting inclusion.
    // This prevents links inbetween urban and surban from remaining rural.
    if GetSelectInclusion() = "Enclosed" then do
        reset_inclusion = "true"
        SetSelectInclusion("Intersecting")
    end

    // Add highway links to map and add AreaType field
    hwy_dbd = hwy_dbd
    {nLayer, llyr} = GetDBLayers(hwy_dbd)
    llyr = AddLayer(map, llyr, hwy_dbd, llyr)
    a_fields = {{"AreaType", "Character", 10, }}
    RunMacro("Add Fields", {view: llyr, a_fields: a_fields})
    SetLayer(llyr)
    SelectByQuery("primary", "several", "Select * where DTWB contains 'D' or DTWB contains 'T'")

    // Loop over each area type starting with most dense.  Skip the first.
    // All remaining links after this loop will be tagged with the lowest
    // area type. Secondary links (walk network) not tagged.
    for t = area_tbl.length to 2 step -1 do
        type = area_tbl[t].AreaType

        // Select TAZs of current type
        SetView(jv)
        query = "Select * where " + se_vw + ".AreaType = '" + type + "'"
        n = SelectByQuery("selection", "Several", query)

        if n > 0 then do
            // Create buffer and add it to the map
            buffer_dbd = GetTempFileName(".dbd")
            opts = null
            opts.Exterior = "Merged"
            opts.Interior = "Merged"
            CreateBuffers(buffer_dbd, "buffer", {"selection"}, "Value", {100/5280}, )
            bLyr = AddLayer(map, "buffer", buffer_dbd, "buffer")

            // Select links within the buffer that haven't been updated already
            SetLayer(llyr)
            n2 = SelectByVicinity(
                "links", "several", taz_lyr + "|selection", 0, 
                {"Source And": "primary"}
            )
            query = "Select * where AreaType <> null"
            n2 = SelectByQuery("links", "Less", query)

            // Remove buffer from map
            DropLayer(map, bLyr)

            if n2 > 0 then do
                // For these links, update their area type
                v_at = Vector(n2, "String", {{"Constant", type}})
                SetDataVector(llyr + "|links", "AreaType", v_at, )
            end
        end
    end

    // Select all remaining links and assign them to the
    // first (lowest density) area type.
    SetLayer(llyr)
    query = "Select * where AreaType = null and (DTWB contains 'D' or DTWB contains 'T')"
    n = SelectByQuery("links", "Several", query)
    if n > 0 then do
        type = area_tbl[1].AreaType
        v_at = Vector(n, "String", {{"Constant", type}})
        SetDataVector(llyr + "|links", "AreaType", v_at, )
    end

    // If this script modified the user setting for inclusion, change it back.
    if reset_inclusion = "true" then SetSelectInclusion("Enclosed")
EndMacro

/*
DC models for residents and visitors need the cluster information on the se
data table.
*/

Macro "Add Cluster Info to SE Data" (Args)

    se_file = Args.DemographicOutputs
    taz_file = Args.TAZGeography

    se = CreateObject("Table", se_file)
    se.AddField({FieldName: "Cluster", type: "integer"})
    se.AddField({FieldName: "ClusterName", type: "string"})
    taz = CreateObject("Table", taz_file)
    join = se.Join({
        Table: taz,
        LeftFields: "TAZ",
        RightFields: "TAZID"
    })
    join.Cluster = join.District7
    join.ClusterName = "c" + String(join.District7)
endmacro

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

macro "Speeds and Capacities" (Args, Result)
    ret_value = 1
    // input data files
    LineDB = Args.HighwayDatabase
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    fields = {
    {FieldName: "ABSpeedLimit"}, 
    {FieldName: "BASpeedLimit"}, 
    {FieldName: "ABFreeFlowSpeed"}, 
    {FieldName: "BAFreeFlowSpeed"}, 
    {FieldName: "ABFreeFlowTime"}, 
    {FieldName: "BAFreeFlowTime"}, 
    {FieldName: "ABAMTime"}, 
    {FieldName: "BAAMTime"}, 
    {FieldName: "ABPMTime"}, 
    {FieldName: "BAPMTime"}, 
    {FieldName: "ABOPTime"}, 
    {FieldName: "BAOPTime"}, 
    {FieldName: "cap_phpl"}, 
    {FieldName: "ABHourlyCapacityAM"}, 
    {FieldName: "BAHourlyCapacityAM"},  
    {FieldName: "ABHourlyCapacityPM"}, 
    {FieldName: "BAHourlyCapacityPM"},  
    {FieldName: "ABHourlyCapacityOP"}, 
    {FieldName: "BAHourlyCapacityOP"},  
    {FieldName: "ABCapacityAM"}, 
    {FieldName: "BACapacityAM"},  
    {FieldName: "ABCapacityPM"}, 
    {FieldName: "BACapacityPM"},  
    {FieldName: "ABCapacityOP"}, 
    {FieldName: "BACapacityOP"},  
    {FieldName: "ABAlpha"}, 
    {FieldName: "BAAlpha"}, 
    {FieldName: "ABBeta"}, 
    {FieldName: "BABeta"},
    {FieldName: "Mode"}
   }
    Line.AddFields({Fields: fields})

    Line.HCMMedian = if Line.HCMType = "MLHighway" or Line.HCMType = "Freeway"
        then "Restrictive"
        else Line.HCMMedian

    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: {"HCMType", "AreaType", "HCMMedian"}, RightFields: {"HCMType", "AreaType", "HCMMedian"}})
    join.cap_phpl = join.cape_phpl
    periods = {"AM", "PM", "OP"}
    for period in periods do
        // hourly capacities
        join.("ABHourlyCapacity" + period) = if join.("AB_LANE" + period) <> null and join.Dir >= 0 
            then join.cape_phpl * join.("AB_LANE" + period) 
            else if join.("AB_LANE" + period) = null and join.Dir = -1 
                then null 
                else join.cape_phpl
        join.("BAHourlyCapacity" + period) = if join.("BA_LANE" + period) <> null and join.Dir <= 0 
            then join.cape_phpl * join.("BA_LANE" + period) 
            else if join.("BA_LANE" + period) = null and join.Dir = 1 
                then null 
                else join.cape_phpl

        // period capacities
        join.("ABCapacity" + period) = join.("ABHourlyCapacity" + period) * Args.(period + "CapFactor")
        join.("BACapacity" + period) = join.("BAHourlyCapacity" + period) * Args.(period + "CapFactor")
    end
    join.ABAlpha = join.Alpha
    join.BAAlpha = join.Alpha
    join.ABBeta = join.Beta
    join.BABeta = join.Beta
    join.ABSpeedLimit = join.PostedSpeed
    join.BASpeedLimit = join.PostedSpeed
    join.ABFreeFlowSpeed = join.PostedSpeed + join.ModifyPosted
    join.BAFreeFlowSpeed = join.PostedSpeed + join.ModifyPosted
    join.ABFreeFlowTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAFreeFlowTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABAMTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAAMTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABPMTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAPMTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.ABOPTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAOPTime = join.Length / join.BAFreeFlowSpeed * 60   
    join.Mode = 1
    join = null

    quit:
    Return(ret_value)


endmacro

/*

*/

macro "CalculateTransitSpeeds Oahu" (Args, Result)
    ret_value = 1
    // input data files
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    fields = {
    {FieldName: "ABWalkTime"}, 
    {FieldName: "BAWalkTime"}, 
    {FieldName: "ABBikeTime"}, 
    {FieldName: "BABikeTime"}, 
    {FieldName: "ABTransitFactor"}, 
    {FieldName: "BATransitFactor"}, 
    {FieldName: "ABTransitSpeed"}, 
    {FieldName: "BATransitSpeed"}, 
    {FieldName: "ABTransitTime"}, 
    {FieldName: "BATransitTime"}, 
    {FieldName: "ABTransitSpeedAM"}, 
    {FieldName: "BATransitSpeedAM"}, 
    {FieldName: "ABTransitTimeAM"}, 
    {FieldName: "BATransitTimeAM"}, 
    {FieldName: "ABTransitSpeedPM"}, 
    {FieldName: "BATransitSpeedPM"}, 
    {FieldName: "ABTransitTimePM"}, 
    {FieldName: "BATransitTimePM"}, 
    {FieldName: "ABTransitSpeedOP"}, 
    {FieldName: "BATransitSpeedOP"}, 
    {FieldName: "ABTransitTimeOP"}, 
    {FieldName: "BATransitTimeOP"} 
   }
    Line.AddFields({Fields: fields})

    // Bike/Walk times
    Line.WalkSpeed = if Line.WalkSpeed = null then Args.WalkSpeed else Line.WalkSpeed
    Line.BikeSpeed = if Line.BikeSpeed = null then Args.BikeSpeed else Line.BikeSpeed
    Line.ABWalkTime = Line.Length / Line.WalkSpeed * 60
    Line.BAWalkTime = Line.Length / Line.WalkSpeed * 60
    Line.ABBikeTime = Line.Length / Line.BikeSpeed * 60
    Line.BABikeTime = Line.Length / Line.BikeSpeed * 60

    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: {"HCMType", "AreaType", "HCMMedian"}, RightFields: {"HCMType", "AreaType", "HCMMedian"}})
    join.ABTransitFactor = join.TransitFactor
    join.BATransitFactor = join.TransitFactor
    join.ABTransitSpeed = join.ABFreeFlowSpeed / join.ABTransitFactor
    join.BATransitSpeed = join.BAFreeFlowSpeed / join.BATransitFactor
    join.ABTransitTime = join.Length / join.ABTransitSpeed * 60
    join.BATransitTime = join.Length / join.BATransitSpeed * 60
    join.ABTransitSpeedAM = join.ABFreeFlowSpeed / join.ABTransitFactor
    join.BATransitSpeedAM = join.BAFreeFlowSpeed / join.BATransitFactor
    join.ABTransitTimeAM = join.Length / join.ABTransitSpeedAM * 60
    join.BATransitTimeAM = join.Length / join.BATransitSpeedAM * 60
    join.ABTransitSpeedPM = join.ABFreeFlowSpeed / join.ABTransitFactor
    join.BATransitSpeedPM = join.BAFreeFlowSpeed / join.BATransitFactor
    join.ABTransitTimePM = join.Length / join.ABTransitSpeedPM * 60
    join.BATransitTimePM = join.Length / join.BATransitSpeedPM * 60
    join.ABTransitSpeedOP = join.ABFreeFlowSpeed / join.ABTransitFactor
    join.BATransitSpeedOP = join.BAFreeFlowSpeed / join.BATransitFactor
    join.ABTransitTimeOP = join.Length / join.ABTransitSpeedOP * 60
    join.BATransitTimeOP = join.Length / join.BATransitSpeedOP * 60
    // Calculate for non-drivable links (e.g. transit only)
    join.SelectByQuery({
        SetName: "transit_only",
        Query: "T = 1 and D = 0"
    })
    periods = {"AM", "PM", "OP"}
    dirs = {"AB", "BA"}
    for period in periods do
        for dir in dirs do
            posted_speed = if join.PostedSpeed = null
                then 25
                else join.PostedSpeed
            join.(dir + "TransitSpeed" + period) = posted_speed
            join.(dir + "TransitTime" + period) = join.Length / posted_speed * 60
            // Fill in the non-period fields just to avoid confusion when
            // looking at the link layer.
            join.(dir + "TransitSpeed") = posted_speed
            join.(dir + "TransitTime") = join.Length / posted_speed * 60
        end
    end

    RS = CreateObject("Table", {FileName: RouteSystem, LayerType: "Route"})
    RouteLayer = RS.GetView()
    STOP = CreateObject("Table", {FileName: RouteSystem, LayerType: "Stop"})
    fields = {
    {FieldName: "NodeID", Type: "integer"}}
    STOP.AddFields({Fields: fields})
    n = TagRouteStopsWithNode(RouteLayer, , "NodeID", 10) // fill stop layer with node ID for each stop
    RS = null
    STOP = null

       quit:
    Return(ret_value)

endmacro

// /*
//     04/13/2023 - modified code such that the three access modes (walk, knr, pnr) become user classes 
//     04/20/2023 - use updated database with mode flag fields T,D,W
//     04/21/2023 - created skimming code
// */

// Macro "test"
//     Args.modelPath = "G:\\USERS\\JIAN\\Kyle\\Oahu_04-21\\data\\"

//     RunMacro("Setup Files", Args)
//     RunMacro("Setup Parameters", Args)

//     RunMacro("Preprocess Network", Args)
//     RunMacro("Create Transit Networks", Args)
//     RunMacro("Transit Skimming", Args)
//     RunMacro("Transit Assignment", Args)

//     ShowMessage("Done")
// endmacro

// Macro "Setup Files" (Args)
//     path = Args.modelPath

//     inPath = path + "input\\"
//     outPath = path + "output\\"

//     // inputs
//         // network
//     Args.lineDB = inPath + "network\\scenario_links.dbd"
//     Args.Routes = inPath + "network\\scenario_routes.rts"
//     Args.TransitModeTable = inPath + "network\\transit_mode_table.bin"

//         // assignment
//     Args.[AM walk OD Matrix] = inPath + "assign\\AM_walk_PA.mtx"
//     Args.[AM pnr OD Matrix] = inPath + "assign\\AM_pnr_PA.mtx"
//     Args.[AM knr OD Matrix] = inPath + "assign\\AM_knr_PA.mtx"

//     // outputs
//         // network
//     Args.[AM walk transit net] = outPath + "network\\AM_walk.tnw"
//     Args.[AM pnr transit net]  = outPath + "network\\AM_pnr.tnw"
//     Args.[AM knr transit net]  = outPath + "network\\AM_knr.tnw"

//         // skim
//     Args.[AM walk bus Skim Matrix] = outPath + "skim\\AM_walk_bus_skim.mtx"
//     Args.[AM walk brt Skim Matrix] = outPath + "skim\\AM_walk_brt_skim.mtx"
//     Args.[AM walk rail Skim Matrix]= outPath + "skim\\AM_walk_rail_skim.mtx"
//     Args.[AM walk all Skim Matrix] = outPath + "skim\\AM_walk_all_skim.mtx"

//     Args.[AM pnr bus Skim Matrix] = outPath + "skim\\AM_pnr_bus_skim.mtx"
//     Args.[AM pnr brt Skim Matrix] = outPath + "skim\\AM_pnr_brt_skim.mtx"
//     Args.[AM pnr rail Skim Matrix]= outPath + "skim\\AM_pnr_rail_skim.mtx"

//     Args.[AM knr bus Skim Matrix] = outPath + "skim\\AM_knr_bus_skim.mtx"
//     Args.[AM knr brt Skim Matrix] = outPath + "skim\\AM_knr_brt_skim.mtx"
//     Args.[AM knr rail Skim Matrix]= outPath + "skim\\AM_knr_rail_skim.mtx"

//         // assignment
//     Args.[AM walk bus LineFlow Table] = outPath + "assign\\AM_walk_bus_LineFlow.bin"
//     Args.[AM walk bus WalkFlow Table] = outPath + "assign\\AM_walk_bus_WalkFlow.bin"
//     Args.[AM walk bus AggrFlow Table] = outPath + "assign\\AM_walk_bus_AggrFlow.bin"
//     Args.[AM walk bus Boarding Table] = outPath + "assign\\AM_walk_bus_Boarding.bin"

//     Args.[AM walk brt LineFlow Table] = outPath + "assign\\AM_walk_brt_LineFlow.bin"
//     Args.[AM walk brt WalkFlow Table] = outPath + "assign\\AM_walk_brt_WalkFlow.bin"
//     Args.[AM walk brt AggrFlow Table] = outPath + "assign\\AM_walk_brt_AggrFlow.bin"
//     Args.[AM walk brt Boarding Table] = outPath + "assign\\AM_walk_brt_Boarding.bin"

//     Args.[AM walk rail LineFlow Table] = outPath + "assign\\AM_walk_rail_LineFlow.bin"
//     Args.[AM walk rail WalkFlow Table] = outPath + "assign\\AM_walk_rail_WalkFlow.bin"
//     Args.[AM walk rail AggrFlow Table] = outPath + "assign\\AM_walk_rail_AggrFlow.bin"
//     Args.[AM walk rail Boarding Table] = outPath + "assign\\AM_walk_rail_Boarding.bin"

//     Args.[AM walk all LineFlow Table] = outPath + "assign\\AM_walk_all_LineFlow.bin"
//     Args.[AM walk all WalkFlow Table] = outPath + "assign\\AM_walk_all_WalkFlow.bin"
//     Args.[AM walk all AggrFlow Table] = outPath + "assign\\AM_walk_all_AggrFlow.bin"
//     Args.[AM walk all Boarding Table] = outPath + "assign\\AM_walk_all_Boarding.bin"

//     Args.[AM pnr bus LineFlow Table] = outPath + "assign\\AM_pnr_bus_LineFlow.bin"
//     Args.[AM pnr bus WalkFlow Table] = outPath + "assign\\AM_pnr_bus_WalkFlow.bin"
//     Args.[AM pnr bus AggrFlow Table] = outPath + "assign\\AM_pnr_bus_AggrFlow.bin"
//     Args.[AM pnr bus Boarding Table] = outPath + "assign\\AM_pnr_bus_Boarding.bin"

//     Args.[AM pnr brt LineFlow Table] = outPath + "assign\\AM_pnr_brt_LineFlow.bin"
//     Args.[AM pnr brt WalkFlow Table] = outPath + "assign\\AM_pnr_brt_WalkFlow.bin"
//     Args.[AM pnr brt AggrFlow Table] = outPath + "assign\\AM_pnr_brt_AggrFlow.bin"
//     Args.[AM pnr brt Boarding Table] = outPath + "assign\\AM_pnr_brt_Boarding.bin"

//     Args.[AM pnr rail LineFlow Table] = outPath + "assign\\AM_pnr_rail_LineFlow.bin"
//     Args.[AM pnr rail WalkFlow Table] = outPath + "assign\\AM_pnr_rail_WalkFlow.bin"
//     Args.[AM pnr rail AggrFlow Table] = outPath + "assign\\AM_pnr_rail_AggrFlow.bin"
//     Args.[AM pnr rail Boarding Table] = outPath + "assign\\AM_pnr_rail_Boarding.bin"

//     Args.[AM knr bus LineFlow Table] = outPath + "assign\\AM_knr_bus_LineFlow.bin"
//     Args.[AM knr bus WalkFlow Table] = outPath + "assign\\AM_knr_bus_WalkFlow.bin"
//     Args.[AM knr bus AggrFlow Table] = outPath + "assign\\AM_knr_bus_AggrFlow.bin"
//     Args.[AM knr bus Boarding Table] = outPath + "assign\\AM_knr_bus_Boarding.bin"

//     Args.[AM knr brt LineFlow Table] = outPath + "assign\\AM_knr_brt_LineFlow.bin"
//     Args.[AM knr brt WalkFlow Table] = outPath + "assign\\AM_knr_brt_WalkFlow.bin"
//     Args.[AM knr brt AggrFlow Table] = outPath + "assign\\AM_knr_brt_AggrFlow.bin"
//     Args.[AM knr brt Boarding Table] = outPath + "assign\\AM_knr_brt_Boarding.bin"

//     Args.[AM knr rail LineFlow Table] = outPath + "assign\\AM_knr_rail_LineFlow.bin"
//     Args.[AM knr rail WalkFlow Table] = outPath + "assign\\AM_knr_rail_WalkFlow.bin"
//     Args.[AM knr rail AggrFlow Table] = outPath + "assign\\AM_knr_rail_AggrFlow.bin"
//     Args.[AM knr rail Boarding Table] = outPath + "assign\\AM_knr_rail_Boarding.bin"

// endmacro

// Macro "Setup Parameters" (Args)

// //jz    Args.Periods = {"EA", "AM", "MD", "PM", "NT"}
//     Args.Periods = {"AM"}   // for testing, only run AM

//     // For testing, the KNR and PNR nodes are the exact same, but they will
//     // be different in the final model, so please create separate classes.
//     Args.AccessModes = {"walk", "pnr", "knr"}
//     Args.TransitModes = {"bus", "brt", "rail", "all"}

//     Args.[Transit Speed] = 25 // mile/hr, default
//     Args.[Drive Speed] = 25   // mile/hr, default
//     Args.[Walk Speed] = 3     // mile/hr

// endmacro

// Macro "Preprocess Network" (Args)

//     lineDB = Args.lineDB
//     Periods = Args.periods
//     defTranSpeed = Args.[Transit Speed]
//     defDrvSpeed = Args.[Drive Speed]
//     walkSpeed = Args.[Walk Speed]    

//     objLyrs = CreateObject("AddDBLayers", {FileName: lineDB})
//     {, linkLyr} = objLyrs.Layers

//         // add Mode field in link table
//     obj = CreateObject("CC.ModifyTableOperation",  linkLyr)
//     obj.FindOrAddField("Mode", "integer")
//     obj.Apply()

//         // compute transit link time
//     SetLayer(linkLyr)
//     tranQry = "Select * Where T = 1"
//     numTrans = SelectByQuery("transit links", "Several", tranQry,)
//     tranVwSet = linkLyr + "|transit links"
//     linkDir = GetDataVector(tranVwSet, "Dir", )
//     linkLen = GetDataVector(tranVwSet, "Length", )
//     postSpeed = GetDataVector(tranVwSet, "PostedSpeed", )
//     tranSpeed = if postSpeed > 0 then postSpeed else defTranSpeed // we may need to apply transit speed factor here
//     tranTime = linkLen / tranSpeed * 60
//     tranTime = max(0.01, tranTime)
//     TranShortModes = {"LB", "EB", "FG"}
//     for period in Periods do
//         for i = 1 to TranShortModes.length do
//             shortMode = TranShortModes[i]
//             abTranTimeFld = "AB" + period + shortMode + "Time"
//             baTranTimeFld = "BA" + period + shortMode + "Time"
//             abTranTime = if linkDir >= 0 then tranTime else null
//             baTranTime = if linkDir <= 0 then tranTime else null
//             SetDataVector(tranVwSet, abTranTimeFld, abTranTime,)
//             SetDataVector(tranVwSet, baTranTimeFld, baTranTime,)
//           end // for i
//       end // for period

//         // set walk link flag and their non-transit mode
//     SetLayer(linkLyr)
//     walkQry = "Select * Where W = 1"
//     numWalks = SelectByQuery("walk links", "Several", walkQry,)
//     walkVwSet = linkLyr + "|walk links"
//     oneVec = Vector(numWalks, "Long", {{"Constant", 1}})
//     SetDataVector(walkVwSet, "Mode", oneVec,) //NOTE: mode ID for non-transit mode is 1

//         // compute walk link time
//     linkLen = GetDataVector(walkVwSet, "Length", )
//     walkTime = linkLen / walkSpeed * 60
//     walkTime = max(0.01, walkTime)
//     SetDataVector(walkVwSet, "WalkTime", walkTime,)

//         // set drive link flag and their non-transit mode
//     SetLayer(linkLyr)
//     driveQry = "Select * Where D = 1"
//     numDrives = SelectByQuery("drive links", "Several", driveQry,)
//     driveVwSet = linkLyr + "|drive links"
//     oneVec = Vector(numDrives, "Long", {{"Constant", 1}})
//     SetDataVector(driveVwSet, "Mode", oneVec,) //NOTE: walk links and drive links may overlap

//         // compute drive link time
//     linkDir = GetDataVector(driveVwSet, "Dir", )
//     linkLen = GetDataVector(driveVwSet, "Length", )
//     postSpeed = GetDataVector(driveVwSet, "PostedSpeed", )
//     drvSpeed = if postSpeed > 0 then postSpeed else defDrvSpeed
//     drvTime = linkLen / drvSpeed * 60
//     drvTime = max(0.01, drvTime)
//     abDrvTime = if linkDir >= 0 then drvTime else null
//     baDrvTime = if linkDir <= 0 then drvTime else null
//     SetDataVector(driveVwSet, "ABDriveTime", abDrvTime,)
//     SetDataVector(driveVwSet, "BADriveTime", baDrvTime,)

// endMacro





// Macro "Transit Skimming" (Args)
//     on error do
//         ShowMessage("Transit Skims " + GetLastError())
//         return()
//       end

//     Periods = Args.periods
//     AccessModes = Args.AccessModes
    
//     for period in Periods do
//         for acceMode in AccessModes do
//             ok = RunMacro("transit skim", Args, period, acceMode)
//             if !ok then goto quit
//           end
//       end

//   quit:
//     return(ok)
// endMacro



// Macro "Transit Assignment" (Args)
//     on error do
//         ShowMessage("Transit Assignment " + GetLastError())
//         return()
//       end

//     Periods = Args.periods
//     AccessModes = Args.AccessModes
    
//     for period in Periods do
//         for acceMode in AccessModes do
//             ok = RunMacro("transit assign", Args, period, acceMode)
//             if !ok then goto quit
//           end
//       end

//   quit:
//     return(ok)
// endmacro

// Macro "transit assign" (Args, period, acceMode)
//     rsFile = Args.Routes
//     tnwFile = Args.(period + " " + acceMode + " Transit Net") // AM_walk.tnw, AM_pnr.tnw, AM_knr.tnw
//     odMatrix = Args.(period + " " + acceMode + " OD Matrix")  // AM_walk_PA.mtx, AM_pnr_PA.mtx, AM_knr_PA.mtx

//     if acceMode = "walk" then 
//         TransModes = Args.TransitModes // {"bus", "brt", "rail", "all"}
//     else // if "pnr" or "knr"
//         TransModes = Subarray(Args.TransitModes, 1, 3) // {"bus", "brt", "rail"}

//     for transMode in TransModes do
//         label = period + " " + acceMode + " " + transMode

//         lineFlowTable = Args.(label + " LineFlow Table")
//         walkFlowTable = Args.(label + " WalkFlow Table")
//         aggrFlowTable = Args.(label + " AggrFlow Table")
//         boardVolTable = Args.(label + " Boarding Table")

//         ok = RunMacro("Set Transit Network", Args, period, acceMode, transMode)
//         if !ok then goto quit

//             // do assignment
//         className = period + "-" + acceMode + "-" + transMode

//         obj = CreateObject("Network.PublicTransportAssignment", {RS: rsFile, NetworkName: tnwFile})

//         obj.ODLayerType = "Node"
//         obj.Method = "PF"
//         obj.AddDemandMatrix({Class: className, Matrix: {MatrixFile: odMatrix, Matrix: "weight", RowIndex: "ProdTAZ", ColumnIndex: "AttrTAZ"}})

//         obj.FlowTable             = lineFlowTable
//         obj.WalkFlowTable         = walkFlowTable
//         obj.TransitLinkFlowsTable = aggrFlowTable
//         obj.OnOffTable            = boardVolTable

//         ok = obj.Run()
//         if !ok then goto quit
//       end // for transMode

//   quit:
//     return(ok)
// endmacro

/*

*/

macro "BuildNetworks Oahu" (Args, Result)

    RunMacro("BuildHighwayNetwork Oahu", Args)
    RunMacro("Create Transit Networks", Args)
    return(1)
EndMacro

/*

*/

macro "BuildHighwayNetwork Oahu" (Args)

    periods = {"AM", "PM", "OP"}
    out_dir = Args.[Output Folder] + "/skims"

    ret_value = 1
    LineDB = Args.HighwayDatabase
    // TurnPenaltyFile = Args.TurnPenalties
    
    for period in periods do
        netfile = out_dir + "/highwaynet_" + period + ".net"
    
        netObj = CreateObject("Network.Create")
        netObj.LayerDB = LineDB
        netObj.Filter =  "D = 1 and (nz(AB_LANE" + period + ") + nz(BA_LANE" + period + ") > 0)" 
        netObj.AddLinkField({Name: "FreeFlowTime", Field: {"ABFreeFlowTime", "BAFreeFlowTime"}})
        netObj.AddLinkField({Name: "Time", Field: {"AB" + period + "Time", "BA" + period + "Time"}})
        // netObj.AddLinkField({Name: "AMTime", Field: {"ABAMTime", "BAAMTime"}})
        // netObj.AddLinkField({Name: "PMTime", Field: {"ABPMTime", "BAPMTime"}})
        // netObj.AddLinkField({Name: "OPTime", Field: {"ABOPTime", "BAOPTime"}})
        // netObj.AddLinkField({Name: "HourlyCapacity", Field: {"ABHourlyCapacity", "BAHourlyCapacity"}})
        // netObj.AddLinkField({Name: "HourlyCapacityAM", Field: {"ABHourlyCapacityAM", "BAHourlyCapacityAM"}})
        // netObj.AddLinkField({Name: "HourlyCapacityPM", Field: {"ABHourlyCapacityPM", "BAHourlyCapacityPM"}})
        // netObj.AddLinkField({Name: "HourlyCapacityOP", Field: {"ABHourlyCapacityOP", "BAHourlyCapacityOP"}})
        netObj.AddLinkField({Name: "Capacity", Field: {"ABCapacity" + period, "BACapacity" + period}})
        // netObj.AddLinkField({Name: "AMCapacity", Field: {"ABCapacityAM", "BACapacityAM"}})
        // netObj.AddLinkField({Name: "PMCapacity", Field: {"ABCapacityPM", "BACapacityPM"}})
        // netObj.AddLinkField({Name: "OPCapacity", Field: {"ABCapacityOP", "BACapacityOP"}})
        netObj.AddLinkField({Name: "WalkTime", Field: {"ABWalkTime", "BAWalkTime"}})
        netObj.AddLinkField({Name: "BikeTime", Field: {"ABBikeTime", "BABikeTime"}})
        netObj.AddLinkField({Name: "Alpha", Field: {"ABAlpha", "BAAlpha"}})
        netObj.AddLinkField({Name: "Beta", Field: {"ABBeta", "BABeta"}})
        netObj.OutNetworkName = netfile
        netObj.Run()
        
        netSetObj = null
        netSetObj = CreateObject("Network.Settings")
        netSetObj.LayerDB = LineDB
        netSetObj.LoadNetwork(netfile)
        netSetObj.CentroidFilter = "Centroid = 1"
            // netSetObj.SetPenalties({LinkPenaltyTable: TurnPenaltyFile, PenaltyField: "Penalty"})
        netSetObj.Run()

    end

    quit:
    Return(ret_value)

endmacro

/*

*/

// macro "BuildTransitNetwork Oahu" (Args)

//     ret_value = 1
//     RouteSystem = Args.TransitRoutes
//     TransitTNW = Args.TransitNetwork

//     netObj = CreateObject("Network.CreatePublic")
//     netObj.LayerRS = RouteSystem
//     netObj.OutNetworkName = TransitTNW
//     netObj.StopToNodeTagField = "NodeID"
//     netObj.IncludeWalkLinks = true
//     netObj.WalkLinkFilter = "W = 1"
//     netObj.IncludeDriveLinks = true
//     netObj.DriveLinkFilter = "D = 1"
//     netObj.AddRouteField({Name: "PeakHeadway", Field: "AMHeadway"})
//     netObj.AddRouteField({Name: "OffpeakHeadway", Field: "MDHeadway"})
//     netObj.AddRouteField({Name: "Fare", Field: "Fare"})
//     netObj.AddLinkField({Name: "TransitTime", TransitFields: {"ABTransitTime", "BATransitTime"}, NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
//     netObj.AddLinkField({Name: "TransitTimeAM", TransitFields: {"ABTransitTimeAM", "BATransitTimeAM"}, NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
//     netObj.AddLinkField({Name: "TransitTimePM", TransitFields: {"ABTransitTimePM", "BATransitTimePM"}, NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
//     netObj.AddLinkField({Name: "TransitTimeOP", TransitFields: {"ABTransitTimeOP", "BATransitTimeOP"}, NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
//     netObj.AddLinkField({Name: "Time", TransitFields: {"ABFreeflowTime", "BAFreeFlowTime"}, NonTransitFields: {"ABFreeflowTime", "BAFreeFlowTime"}})
//     netObj.AddLinkField({Name: "AMTime", TransitFields: {"ABAMTime", "BAAMTime"}, NonTransitFields: {"ABAMTime", "BAAMTime"}})
//     netObj.AddLinkField({Name: "PMTime", TransitFields: {"ABPMTime", "BAPMTime"}, NonTransitFields: {"ABPMTime", "BAPMTime"}})
//     netObj.AddLinkField({Name: "OPTime", TransitFields: {"ABOPTime", "BAOPTime"}, NonTransitFields: {"ABOPTime", "BAOPTime"}})
//     netObj.Run()

//        quit:
//     Return(ret_value)

// endmacro

Macro "Create Transit Networks" (Args)

    rsFile = Args.TransitRoutes
    Periods = {"AM", "PM", "OP"}
    AccessModes = Args.AccessModes
    skim_dir = Args.OutputSkims

    objLyrs = CreateObject("AddRSLayers", {FileName: rsFile})
    rtLyr = objLyrs.RouteLayer

    // Retag stops to nodes. While this step is done by the route manager
    // during scenario creation, a user might create a new route to test after
    // creating the scenario. This makes sure it 'just works'.
    TagRouteStopsWithNode(rtLyr,, "Node_ID", 0.2)

    for period in Periods do
        for acceMode in AccessModes do
            tnwFile = skim_dir + "\\transit\\" + period + "_" + acceMode + ".tnw"
            o = CreateObject("Network.CreateTransit")
            o.LayerRS = rsFile
            o.OutNetworkName = tnwFile
            o.UseModes({TransitModeField: "Mode", NonTransitModeField: "Mode"})

            // route attributes
            o.RouteFilter = period + "Headway > 0 & Mode > 0"
            o.AddRouteField({Name: period + "Headway", Field: period + "Headway"})
            o.AddRouteField({Name: "Fare", Field: "Fare"})

            // stop attributes
            o.StopToNodeTagField = "Node_ID"

            // link attributes
            o.AddLinkField({Name: "bus_time", TransitFields: {"ABTransitTime" + period, "BATransitTime" + period},
                                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
            // rail time never changes so just use AM
            o.AddLinkField({Name: "rail_time", TransitFields: {"ABTransitTimeAM", "ABTransitTimeAM"},
                                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})

            // drive attributes
            o.IncludeDriveLinks = true
            o.DriveLinkFilter = "D = 1"
            o.AddLinkField({Name: "DriveTime",
                            TransitFields:    {"AB" + period + "Time", "BA" + period + "Time"}, 
                            NonTransitFields: {"AB" + period + "Time", "BA" + period + "Time"}})

            // walk attributes
            o.IncludeWalkLinks = true
            o.WalkLinkFilter = "W = 1"
            o.AddLinkField({Name:             "WalkTime",
                            TransitFields:    {"ABWalkTime", "BAWalkTime"}, 
                            NonTransitFields: {"ABWalkTime", "BAWalkTime"}})

            ok = o.Run()

            RunMacro("Set Transit Network", Args, period, acceMode,)
          end // for acceMode
      end // for period
endMacro

Macro "Set Transit Network" (Args, period, acceMode, currTransMode)
    rsFile = Args.TransitRoutes
    modeTable = Args.TransitModeTable
    skim_dir = Args.OutputSkims
    tnwFile = skim_dir + "\\transit\\" + period + "_" + acceMode + ".tnw"

    o = CreateObject("Network.SetPublicPathFinder", {RS: rsFile, NetworkName: tnwFile})

    // define user classes
    UserClasses = null
    ModeUseFld = null
    DrvTimeFld = null
    DrvInUse   = null
    PermitAllW = null
    AllowWacc = null
    ParkFilter = null

    // build class name list and class-specific PnR/KnR option array
    transit_modes = RunMacro("Get Transit Net Def Col Names", modeTable)
    if acceMode = "w" 
        then TransModes = transit_modes
        // if "pnr" or "knr" remove 'all'
        else TransModes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)

    for transMode in TransModes do
        UserClasses = UserClasses + {period + "-" + acceMode + "-" + transMode}
        ModeUseFld = ModeUseFld + {transMode}
        PermitAllW = PermitAllW + {false}

        if acceMode = "w" then do
            DrvTimeFld = DrvTimeFld + {}
            DrvInUse = DrvInUse + {false}
            AllowWacc = AllowWacc + {true}
            ParkFilter = ParkFilter + {}
        end else do
            DrvTimeFld = DrvTimeFld + {"DriveTime"}
            DrvInUse = DrvInUse + {true}
            AllowWacc = AllowWacc + {false}

            if acceMode = "knr" 
                then ParkFilter = ParkFilter + {"KNR = 1"}
                else ParkFilter = ParkFilter + {"PNR = 1"}
        end // else (if acceMode)
    end // for transMode

    o.UserClasses = UserClasses

    o.DriveTime = DrvTimeFld
    DrvOpts = null
    DrvOpts.InUse = DrvInUse
    DrvOpts.PermitAllWalk = PermitAllW
    DrvOpts.AllowWalkAccess = AllowWacc
    DrvOpts.ParkingNodes = ParkFilter
    if period = "PM" then
        o.DriveEgress(DrvOpts)
    else
        o.DriveAccess(DrvOpts)

    o.CentroidFilter = "Centroid = 1"
    o.LinkImpedance = "bus_time" // default

    o.Parameters(
        {MaxTripCost : 240,
         MaxTransfers: 2,
         VOT         : 0.1984 // $/min (40% of the median wage)
        })

    o.AccessControl(
        {PermitWalkOnly:     false,
         MaxWalkAccessPaths: 10
        })

    o.Combination(
        {CombinationFactor: .1
        })

    o.TimeGlobals(
        {Headway:         14,
         InitialPenalty:  0,
         TransferPenalty: 5,
         MaxInitialWait:  30,
         MaxTransferWait: 10,
         MinInitialWait:  2,
         MinTransferWait: 5,
         Layover:         5, 
         MaxAccessWalk:   45,
         MaxEgressWalk:   45,
         MaxModalTotal:   120
        })

    o.RouteTimeFields(
        {Headway: period + "Headway"
        })

    o.ModeTable(
        {TableName: modeTable,
        // A field in the mode table that contains a list of
        // link network field names. These network field names
        // in turn point to the AB/BA fields on the link layer.
         TimeByMode:          "IVTT",
         ModesUsedField:      ModeUseFld,
         OnlyCombineSameMode: true,
         FreeTransfers:       2
        })

    o.RouteWeights(
        {
         Fare: null,
         Time: null,
         InitialPenalty: null,
         TransferPenalty: null,
         InitialWait: null,
         TransferWeight: null,
         Dwelling: null
        })

    o.GlobalWeights(
        {Fare:            1.0,
         Time:            1.0,
         InitialPenalty:  1.0,
         TransferPenalty: 3.0,
         InitialWait:     3.0,
         TransferWait:    3.0,
         Dwelling:        2.0,
         WalkTimeFactor:  3.0,
         DriveTimeFactor: 1.0
        })

    o.Fare(
        {Type:              "Flat",
         RouteFareField:    "Fare",
         RouteXFareField:   "Fare",
         FareValue:         0.0,
         TransferFareValue: 0.0
        })

    if currTransMode <> null then
        o.CurrentClass = period + "-" + acceMode + "-" + currTransMode 

    ok = o.Run()
    if !ok then goto quit

  quit:
    return(ok)
endMacro