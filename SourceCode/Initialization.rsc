macro "PreprocessData" (Args, Result)
    ret_value = RunMacro("CopyDataToOutputFolder", Args)
    if !ret_value then goto quit

    ret_value = RunMacro("CalculateHighwaySpeedsAndCapacities", Args)
    if !ret_value then goto quit

    ret_value = RunMacro("CalculateTransitSpeeds", Args)
    if !ret_value then goto quit

    ret_value = RunMacro("Compute Intrazonal Matrix", Args)
    if !ret_value then goto quit
  
  quit:
    Return(ret_value)
endmacro


// copies highway network, demographics, and transit routes to output folder for further processing
macro "CopyDataToOutputFolder" (Args)
    ret_value = 1
    hwyinputdb = Args.HighwayInputDatabase 
    hwyoutputdb = Args.HighwayDatabase 
    DemoMaster = Args.Demographics
    DemoOut = Args.DemographicOutputs
    rs_filemaster = Args.TransitRouteInputs
    rs_file = Args.TransitRoutes
    CopyDatabase(hwyinputdb, hwyoutputdb)   // copy master network to copy
    CopyTableFiles(null, "FFB", DemoMaster, , DemoOut, )    // copy master demographics to copy
    tab = CreateObject("Table", DemoOut)
    dropflds = {"EMP_NAICS_11", "EMP_NAICS_21", "EMP_NAICS_22", "EMP_NAICS_23", "EMP_NAICS_31-33", "EMP_NAICS_42", "EMP_NAICS_44-45", "EMP_NAICS_48-49", "EMP_NAICS_51", 
                "EMP_NAICS_52", "EMP_NAICS_53", "EMP_NAICS_54", "EMP_NAICS_55", "EMP_NAICS_56", "EMP_NAICS_61", "EMP_NAICS_62", "EMP_NAICS_71", "EMP_NAICS_72", 
                "EMP_NAICS_81", "EMP_NAICS_92"}
    tab.DropFields({FieldNames: dropflds})
        // add master transit network

    o = CreateObject("DataManager")
    o.AddDataSource("RS", {FileName: rs_filemaster, DataType: "RS"})
    routeLayers = o.GetRouteLayers("RS")

    RouteLayer = routeLayers.RouteLayer
    StopLayer = routeLayers.StopLayer
    LineLayer = routeLayers.LineLayer
    NodeLayer = routeLayers.NodeLayer

    SetLayer(RouteLayer)
    
    args = null
    args.RouteLayer       = RouteLayer
    args.RouteSet         = null
    args.StopSet          = null
    args.TaggedField      = "NodeID"
    args.RepeatedSetName  = "Repeated Stop Tags"
    args.NotTaggedSetName = "Not Tagged"
    args.CreateLog        = true
    // check for repeated tagged nodes and flag if found
    TagInfo = RunMacro("Select Repeated Tagged Nodes", args)
    
    if TagInfo.errorMessage <> null then Throw(GetLastError())
    // create copy of route system

    o.CopyRouteSystem("RS", {TargetRS: rs_file, Settings: {Geography: hwyoutputdb}})
    Return(ret_value)
endmacro

macro "CalculateHighwaySpeedsAndCapacities" (Args, Result)
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
    {FieldName: "ABHourlyCapacity"}, 
    {FieldName: "BAHourlyCapacity"}, 
    {FieldName: "ABAlpha"}, 
    {FieldName: "BAAlpha"}, 
    {FieldName: "ABBeta"}, 
    {FieldName: "BABeta"}
   }
    Line.AddFields({Fields: fields})

    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: "Class", RightFields: "RoadClassName"})
    join.ABHourlyCapacity = if join.AB_NumThruLanes <> null and join.Dir >= 0 then join.Capacity * join.AB_NumThruLanes else if join.AB_NumThruLanes = null and join.Dir = -1 then null else join.Capacity
    join.BAHourlyCapacity = if join.BA_NumThruLanes <> null and join.Dir <= 0 then join.Capacity * join.BA_NumThruLanes else if join.BA_NumThruLanes = null and join.Dir = 1 then null else join.Capacity
    join.ABAlpha = join.Alpha
    join.BAAlpha = join.Alpha
    join.ABBeta = join.Beta
    join.BABeta = join.Beta
    join.ABSpeedLimit = join.SpeedLimit
    join.BASpeedLimit = join.SpeedLimit
    join.ABFreeFlowSpeed = join.FreeFlowSpeed
    join.BAFreeFlowSpeed = join.FreeFlowSpeed
    join.ABFreeFlowTime = join.Length / join.ABFreeFlowSpeed * 60   
    join.BAFreeFlowTime = join.Length / join.BAFreeFlowSpeed * 60   
    join = null

    quit:
    Return(ret_value)


endmacro

macro "CalculateTransitSpeeds" (Args, Result)
    ret_value = 1
    // input data files
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    WalkSpeed = Args.WalkSpeed
    BikeSpeed = Args.BikeSpeed
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
    {FieldName: "BATransitTime"} 
   }
    Line.AddFields({Fields: fields})
    Line.ABWalkTime = Line.Length / WalkSpeed * 60
    Line.BAWalkTime = Line.Length / WalkSpeed * 60
    Line.ABBikeTime = Line.Length / BikeSpeed * 60
    Line.BABikeTime = Line.Length / BikeSpeed * 60

    SpeedCap = Args.SpeedCapacityLookup
    SC = CreateObject("Table", SpeedCap)
    join = Line.Join({Table: SC, LeftFields: "Class", RightFields: "RoadClassName"})
    join.ABTransitFactor = join.TransitFactor
    join.BATransitFactor = join.TransitFactor
    join.ABTransitSpeed = join.ABFreeFlowSpeed / join.ABTransitFactor
    join.BATransitSpeed = join.BAFreeFlowSpeed / join.BATransitFactor
    join.ABTransitTime = join.Length / join.ABTransitSpeed * 60
    join.BATransitTime = join.Length / join.BATransitSpeed * 60

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

macro "BuildNetworks" (Args, Result)

    ret_value = 1
    ret_value = RunMacro("BuildHighwayNetwork", Args)
    if !ret_value then goto quit
    ret_value = RunMacro("BuildTransitNetwork", Args)
    if !ret_value then goto quit
    quit:
    Return(ret_value)
EndMacro

macro "BuildHighwayNetwork" (Args)

    ret_value = 1
    LineDB = Args.HighwayDatabase
    netfile = Args.HighwayNetwork
    TurnPenaltyFile = Args.TurnPenalties
    netObj = CreateObject("Network.Create")
    netObj.LayerDB = LineDB
    netObj.Filter =  "Class <> 'Rail'" 
    netObj.AddLinkField({Name: "FreeFlowTime", Field: {"ABFreeFlowTime", "BAFreeFlowTime"}})
    netObj.AddLinkField({Name: "HourlyCapacity", Field: {"ABHourlyCapacity", "BAHourlyCapacity"}})
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
    netSetObj.CentroidFilter = "Centroid <> null"
        netSetObj.SetPenalties({LinkPenaltyTable: TurnPenaltyFile, PenaltyField: "Penalty"})
    netSetObj.Run()

    quit:
    Return(ret_value)

endmacro

macro "BuildTransitNetwork" (Args)

    ret_value = 1
    RouteSystem = Args.TransitRoutes
    TransitTNW = Args.TransitNetwork

    netObj = CreateObject("Network.CreatePublic")
    netObj.LayerRS = RouteSystem
    netObj.OutNetworkName = TransitTNW
    netObj.StopToNodeTagField = "NodeID"
    netObj.IncludeWalkLinks = true
    netObj.WalkLinkFilter = "Class <> 'Freeway' and Class <> 'Ramp' and Class <> 'System Ramp' and Class <> 'Expressway' and Class <> 'Rail' and Class <> 'Tunnel' and Class <> 'Interstate'"
    netObj.IncludeDriveLinks = true
    netObj.DriveLinkFilter = "Class <> 'Rail'"
    netObj.AddRouteField({Name: "PeakHeadway", Field: "PeakHeadway"})
    netObj.AddRouteField({Name: "OffpeakHeadway", Field: "OffpeakHeadway"})
    netObj.AddRouteField({Name: "Fare", Field: "Fare"})
    netObj.AddLinkField({Name: "TransitTime", TransitFields: {"ABTransitTime", "BATransitTime"}, NonTransitFields: {"ABWalkTime", "BAWalkTime"}})
    netObj.AddLinkField({Name: "Time", TransitFields: {"ABFreeflowTime", "BAFreeFlowTime"}, NonTransitFields: {"ABFreeflowTime", "BAFreeFlowTime"}})
    netObj.Run()

       quit:
    Return(ret_value)

endmacro


    