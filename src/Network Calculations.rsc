/*
    04/13/2023 - modified code such that the three access modes (walk, knr, pnr) become user classes 
    04/20/2023 - use updated database with mode flag fields T,D,W
    04/21/2023 - created skimming code
*/

Macro "test"
    Args.modelPath = "G:\\USERS\\JIAN\\Kyle\\Oahu_04-21\\data\\"

    RunMacro("Setup Files", Args)
    RunMacro("Setup Parameters", Args)

    RunMacro("Preprocess Network", Args)
    RunMacro("Create Transit Networks", Args)
    RunMacro("Transit Skimming", Args)
    RunMacro("Transit Assignment", Args)

    ShowMessage("Done")
endmacro

Macro "Setup Files" (Args)
    path = Args.modelPath

    inPath = path + "input\\"
    outPath = path + "output\\"

    // inputs
        // network
    Args.lineDB = inPath + "network\\scenario_links.dbd"
    Args.Routes = inPath + "network\\scenario_routes.rts"
    Args.TransModeTable = inPath + "network\\transit_mode_table.bin"

        // assignment
    Args.[AM walk OD Matrix] = inPath + "assign\\AM_walk_PA.mtx"
    Args.[AM pnr OD Matrix] = inPath + "assign\\AM_pnr_PA.mtx"
    Args.[AM knr OD Matrix] = inPath + "assign\\AM_knr_PA.mtx"

    // outputs
        // network
    Args.[AM walk transit net] = outPath + "network\\AM_walk.tnw"
    Args.[AM pnr transit net]  = outPath + "network\\AM_pnr.tnw"
    Args.[AM knr transit net]  = outPath + "network\\AM_knr.tnw"

        // skim
    Args.[AM walk bus Skim Matrix] = outPath + "skim\\AM_walk_bus_skim.mtx"
    Args.[AM walk brt Skim Matrix] = outPath + "skim\\AM_walk_brt_skim.mtx"
    Args.[AM walk rail Skim Matrix]= outPath + "skim\\AM_walk_rail_skim.mtx"
    Args.[AM walk all Skim Matrix] = outPath + "skim\\AM_walk_all_skim.mtx"

    Args.[AM pnr bus Skim Matrix] = outPath + "skim\\AM_pnr_bus_skim.mtx"
    Args.[AM pnr brt Skim Matrix] = outPath + "skim\\AM_pnr_brt_skim.mtx"
    Args.[AM pnr rail Skim Matrix]= outPath + "skim\\AM_pnr_rail_skim.mtx"

    Args.[AM knr bus Skim Matrix] = outPath + "skim\\AM_knr_bus_skim.mtx"
    Args.[AM knr brt Skim Matrix] = outPath + "skim\\AM_knr_brt_skim.mtx"
    Args.[AM knr rail Skim Matrix]= outPath + "skim\\AM_knr_rail_skim.mtx"

        // assignment
    Args.[AM walk bus LineFlow Table] = outPath + "assign\\AM_walk_bus_LineFlow.bin"
    Args.[AM walk bus WalkFlow Table] = outPath + "assign\\AM_walk_bus_WalkFlow.bin"
    Args.[AM walk bus AggrFlow Table] = outPath + "assign\\AM_walk_bus_AggrFlow.bin"
    Args.[AM walk bus Boarding Table] = outPath + "assign\\AM_walk_bus_Boarding.bin"

    Args.[AM walk brt LineFlow Table] = outPath + "assign\\AM_walk_brt_LineFlow.bin"
    Args.[AM walk brt WalkFlow Table] = outPath + "assign\\AM_walk_brt_WalkFlow.bin"
    Args.[AM walk brt AggrFlow Table] = outPath + "assign\\AM_walk_brt_AggrFlow.bin"
    Args.[AM walk brt Boarding Table] = outPath + "assign\\AM_walk_brt_Boarding.bin"

    Args.[AM walk rail LineFlow Table] = outPath + "assign\\AM_walk_rail_LineFlow.bin"
    Args.[AM walk rail WalkFlow Table] = outPath + "assign\\AM_walk_rail_WalkFlow.bin"
    Args.[AM walk rail AggrFlow Table] = outPath + "assign\\AM_walk_rail_AggrFlow.bin"
    Args.[AM walk rail Boarding Table] = outPath + "assign\\AM_walk_rail_Boarding.bin"

    Args.[AM walk all LineFlow Table] = outPath + "assign\\AM_walk_all_LineFlow.bin"
    Args.[AM walk all WalkFlow Table] = outPath + "assign\\AM_walk_all_WalkFlow.bin"
    Args.[AM walk all AggrFlow Table] = outPath + "assign\\AM_walk_all_AggrFlow.bin"
    Args.[AM walk all Boarding Table] = outPath + "assign\\AM_walk_all_Boarding.bin"

    Args.[AM pnr bus LineFlow Table] = outPath + "assign\\AM_pnr_bus_LineFlow.bin"
    Args.[AM pnr bus WalkFlow Table] = outPath + "assign\\AM_pnr_bus_WalkFlow.bin"
    Args.[AM pnr bus AggrFlow Table] = outPath + "assign\\AM_pnr_bus_AggrFlow.bin"
    Args.[AM pnr bus Boarding Table] = outPath + "assign\\AM_pnr_bus_Boarding.bin"

    Args.[AM pnr brt LineFlow Table] = outPath + "assign\\AM_pnr_brt_LineFlow.bin"
    Args.[AM pnr brt WalkFlow Table] = outPath + "assign\\AM_pnr_brt_WalkFlow.bin"
    Args.[AM pnr brt AggrFlow Table] = outPath + "assign\\AM_pnr_brt_AggrFlow.bin"
    Args.[AM pnr brt Boarding Table] = outPath + "assign\\AM_pnr_brt_Boarding.bin"

    Args.[AM pnr rail LineFlow Table] = outPath + "assign\\AM_pnr_rail_LineFlow.bin"
    Args.[AM pnr rail WalkFlow Table] = outPath + "assign\\AM_pnr_rail_WalkFlow.bin"
    Args.[AM pnr rail AggrFlow Table] = outPath + "assign\\AM_pnr_rail_AggrFlow.bin"
    Args.[AM pnr rail Boarding Table] = outPath + "assign\\AM_pnr_rail_Boarding.bin"

    Args.[AM knr bus LineFlow Table] = outPath + "assign\\AM_knr_bus_LineFlow.bin"
    Args.[AM knr bus WalkFlow Table] = outPath + "assign\\AM_knr_bus_WalkFlow.bin"
    Args.[AM knr bus AggrFlow Table] = outPath + "assign\\AM_knr_bus_AggrFlow.bin"
    Args.[AM knr bus Boarding Table] = outPath + "assign\\AM_knr_bus_Boarding.bin"

    Args.[AM knr brt LineFlow Table] = outPath + "assign\\AM_knr_brt_LineFlow.bin"
    Args.[AM knr brt WalkFlow Table] = outPath + "assign\\AM_knr_brt_WalkFlow.bin"
    Args.[AM knr brt AggrFlow Table] = outPath + "assign\\AM_knr_brt_AggrFlow.bin"
    Args.[AM knr brt Boarding Table] = outPath + "assign\\AM_knr_brt_Boarding.bin"

    Args.[AM knr rail LineFlow Table] = outPath + "assign\\AM_knr_rail_LineFlow.bin"
    Args.[AM knr rail WalkFlow Table] = outPath + "assign\\AM_knr_rail_WalkFlow.bin"
    Args.[AM knr rail AggrFlow Table] = outPath + "assign\\AM_knr_rail_AggrFlow.bin"
    Args.[AM knr rail Boarding Table] = outPath + "assign\\AM_knr_rail_Boarding.bin"

endmacro

Macro "Setup Parameters" (Args)

//jz    Args.Periods = {"EA", "AM", "MD", "PM", "NT"}
    Args.Periods = {"AM"}   // for testing, only run AM

    // For testing, the KNR and PNR nodes are the exact same, but they will
    // be different in the final model, so please create separate classes.
    Args.AccessModes = {"walk", "pnr", "knr"}
    Args.TransModes = {"bus", "brt", "rail", "all"}

    Args.[Transit Speed] = 25 // mile/hr, default
    Args.[Drive Speed] = 25   // mile/hr, default
    Args.[Walk Speed] = 3     // mile/hr

endmacro

Macro "Preprocess Network" (Args)

    lineDB = Args.lineDB
    Periods = Args.periods
    defTranSpeed = Args.[Transit Speed]
    defDrvSpeed = Args.[Drive Speed]
    walkSpeed = Args.[Walk Speed]    

    objLyrs = CreateObject("AddDBLayers", {FileName: lineDB})
    {, linkLyr} = objLyrs.Layers

        // add Mode field in link table
    obj = CreateObject("CC.ModifyTableOperation",  linkLyr)
    obj.FindOrAddField("Mode", "integer")
    obj.Apply()

        // compute transit link time
    SetLayer(linkLyr)
    tranQry = "Select * Where T = 1"
    numTrans = SelectByQuery("transit links", "Several", tranQry,)
    tranVwSet = linkLyr + "|transit links"
    linkDir = GetDataVector(tranVwSet, "Dir", )
    linkLen = GetDataVector(tranVwSet, "Length", )
    postSpeed = GetDataVector(tranVwSet, "PostedSpeed", )
    tranSpeed = if postSpeed > 0 then postSpeed else defTranSpeed // we may need to apply transit speed factor here
    tranTime = linkLen / tranSpeed * 60
    tranTime = max(0.01, tranTime)
    TranShortModes = {"LB", "EB", "FG"}
    for period in Periods do
        for i = 1 to TranShortModes.length do
            shortMode = TranShortModes[i]
            abTranTimeFld = "AB" + period + shortMode + "Time"
            baTranTimeFld = "BA" + period + shortMode + "Time"
            abTranTime = if linkDir >= 0 then tranTime else null
            baTranTime = if linkDir <= 0 then tranTime else null
            SetDataVector(tranVwSet, abTranTimeFld, abTranTime,)
            SetDataVector(tranVwSet, baTranTimeFld, baTranTime,)
          end // for i
      end // for period

        // set walk link flag and their non-transit mode
    SetLayer(linkLyr)
    walkQry = "Select * Where W = 1"
    numWalks = SelectByQuery("walk links", "Several", walkQry,)
    walkVwSet = linkLyr + "|walk links"
    oneVec = Vector(numWalks, "Long", {{"Constant", 1}})
    SetDataVector(walkVwSet, "Mode", oneVec,) //NOTE: mode ID for non-transit mode is 1

        // compute walk link time
    linkLen = GetDataVector(walkVwSet, "Length", )
    walkTime = linkLen / walkSpeed * 60
    walkTime = max(0.01, walkTime)
    SetDataVector(walkVwSet, "WalkTime", walkTime,)

        // set drive link flag and their non-transit mode
    SetLayer(linkLyr)
    driveQry = "Select * Where D = 1"
    numDrives = SelectByQuery("drive links", "Several", driveQry,)
    driveVwSet = linkLyr + "|drive links"
    oneVec = Vector(numDrives, "Long", {{"Constant", 1}})
    SetDataVector(driveVwSet, "Mode", oneVec,) //NOTE: walk links and drive links may overlap

        // compute drive link time
    linkDir = GetDataVector(driveVwSet, "Dir", )
    linkLen = GetDataVector(driveVwSet, "Length", )
    postSpeed = GetDataVector(driveVwSet, "PostedSpeed", )
    drvSpeed = if postSpeed > 0 then postSpeed else defDrvSpeed
    drvTime = linkLen / drvSpeed * 60
    drvTime = max(0.01, drvTime)
    abDrvTime = if linkDir >= 0 then drvTime else null
    baDrvTime = if linkDir <= 0 then drvTime else null
    SetDataVector(driveVwSet, "ABDriveTime", abDrvTime,)
    SetDataVector(driveVwSet, "BADriveTime", baDrvTime,)

endMacro

Macro "Create Transit Networks" (Args)

    rsFile = Args.Routes
    Periods = Args.periods
    AccessModes = Args.AccessModes

    objLyrs = CreateObject("AddRSLayers", {FileName: rsFile})
    rtLyr = objLyrs.RouteLayer

    // Retag stops to nodes. While this step is done by the route manager
    // during scenario creation, a user might create a new route to test after
    // creating the scenario. This makes sure it 'just works'.
    TagRouteStopsWithNode(rtLyr,, "Node_ID", 0.2)

    for period in Periods do
        for acceMode in AccessModes do
            tnwFile = Args.(period + " " + acceMode + " transit net") // AM_walk.tnw, AM_pnr.tnw, AM_knr.tnw
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
            abLBTimeFld = "AB" + period + "LBTime"
            baLBTimeFld = "BA" + period + "LBTime"
            abEBTimeFld = "AB" + period + "EBTime"
            baEBTimeFld = "BA" + period + "EBTime"
            abFGTimeFld = "AB" + period + "FGTime"
            baFGTimeFld = "BA" + period + "FGTime"
            o.AddLinkField({Name: "LBTime", TransitFields: {abLBTimeFld, baLBTimeFld},
                                            NonTransitFields: "WalkTime"})
            o.AddLinkField({Name: "EBTime", TransitFields: {abEBTimeFld, baEBTimeFld},
                                            NonTransitFields: "WalkTime"})
            o.AddLinkField({Name: "FGTime", TransitFields: {abFGTimeFld, baFGTimeFld},
                                            NonTransitFields: "WalkTime"})

                // drive attributes
            abDrvTimeFld = "ABDriveTime"
            baDrvTimeFld = "BADriveTime"
            o.IncludeDriveLinks = true
            o.DriveLinkFilter = "D = 1"
            o.AddLinkField({Name: "DriveTime",
                            TransitFields:    {"ABDriveTime", "BADriveTime"}, 
                            NonTransitFields: {"ABDriveTime", "BADriveTime"}})

                // walk attributes
            o.IncludeWalkLinks = true
            o.WalkLinkFilter = "W = 1"
            o.AddLinkField({Name:             "WalkTime",
                            TransitFields:    "WalkTime", 
                            NonTransitFields: "WalkTime"})

            ok = o.Run()
            if !ok then goto quit

            RunMacro("Set Transit Network", Args, period, acceMode,)
          end // for acceMode
      end // for period

  quit:
    return(ok)
endMacro

Macro "Set Transit Network" (Args, period, acceMode, currTransMode)
    rsFile = Args.Routes
    modeTable = Args.TransModeTable
    tnwFile = Args.(period + " " + acceMode + " Transit Net") // AM_walk.tnw, AM_pnr.tnw, AM_knr.tnw

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
    if acceMode = "walk" then 
        TransModes = Args.TransModes // {"bus", "brt", "rail", "all"}
    else // if "pnr" or "knr"
        TransModes = Subarray(Args.TransModes, 1, 3) // {"bus", "brt", "rail"}
    for transMode in TransModes do
        UserClasses = UserClasses + {period + "-" + acceMode + "-" + transMode}
        ModeUseFld = ModeUseFld + {transMode}
        PermitAllW = PermitAllW + {true}

        if acceMode = "walk" then do
            DrvTimeFld = DrvTimeFld + {}
            DrvInUse = DrvInUse + {false}
            AllowWacc = AllowWacc + {true}
            ParkFilter = ParkFilter + {}
          end
        else do
            DrvTimeFld = DrvTimeFld + {"DriveTime"}
            DrvInUse = DrvInUse + {true}
            AllowWacc = AllowWacc + {false}

            if acceMode = "knr" then
                ParkFilter = ParkFilter + {"KNR = 1"}
            else
                ParkFilter = ParkFilter + {"PNR = 1"}
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
        o.DriveAccess(DrvOpts)  // temporarily commented out PnR/KnR setting due to bug in Transit API

    o.CentroidFilter = "Centroid = 1"
    o.LinkImpedance = "LBTime" // default

    o.Parameters(
        {MaxTripCost : 240,
         MaxTransfers: 1,
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
         FreeTransfers:       0
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

Macro "Transit Skimming" (Args)
    on error do
        ShowMessage("Transit Skims " + GetLastError())
        return()
      end

    Periods = Args.periods
    AccessModes = Args.AccessModes
    
    for period in Periods do
        for acceMode in AccessModes do
            ok = RunMacro("transit skim", Args, period, acceMode)
            if !ok then goto quit
          end
      end

  quit:
    return(ok)
endMacro

Macro "transit skim" (Args, period, acceMode)
    rsFile = Args.Routes
    tnwFile = Args.(period + " " + acceMode + " Transit Net") // AM_walk.tnw, AM_pnr.tnw, AM_knr.tnw

    if acceMode = "walk" then 
        TransModes = Args.TransModes // {"bus", "brt", "rail", "all"}
    else // if "pnr" or "knr"
        TransModes = Subarray(Args.TransModes, 1, 3) // {"bus", "brt", "rail"}

    for transMode in TransModes do
        label = period + " " + acceMode + " " + transMode + " Skim Matrix"
        outFile = Args.(label)

        ok = RunMacro("Set Transit Network", Args, period, acceMode, transMode)
        if !ok then goto quit

            // do skim
        obj = CreateObject("Network.PublicTransportSkims")

        obj.Network = tnwFile
        obj.LayerRS = rsFile
        obj.Method = "PF"
        obj.SkimByNodes = True
        obj.OriginFilter = "Centroid=1"
        obj.DestinationFilter = "Centroid=1"
//xz        obj.NumberofThreads = 16

        obj.SkimVariables = {"Generalized Cost", "Fare",
                             "In-Vehicle Time",
                             "Initial Wait Time",
                             "Transfer Wait Time",
                             "Initial Penalty Time",
                             "Transfer Penalty Time",
                             "Transfer Walk Time",
                             "Access Walk Time",
                             "Egress Walk Time",
                             "Access Drive Time",
                             "Egress Drive Time",
                             "Dwelling Time",
                             "Total Time",
                             "In-Vehicle Cost",
                             "Initial Wait Cost",
                             "Transfer Wait Cost",
                             "Initial Penalty Cost",
                             "Transfer Penalty Cost",
                             "Transfer Walk Cost",
                             "Access Walk Cost",
                             "Egress Walk Cost",
                             "Access Drive Cost",
                             "Egress Drive Cost",
                             "Dwelling Cost",
                             "Number of Transfers",
                             "In-Vehicle Distance",
                             "Access Drive Distance",
                             "Egress Drive Distance",
                             "Length",
                             "LBTime",
                             "EBTime",
                             "FGTime",
                             "DriveTime",
                             "WalkTime"
                            }
        obj.OutputMatrix({MatrixFile: outFile, MatrixLabel: label, Compression: True})

        ok = obj.Run()
        if !ok then goto quit
      end // for transMode

  quit:
    return(ok)
endmacro

Macro "Transit Assignment" (Args)
    on error do
        ShowMessage("Transit Assignment " + GetLastError())
        return()
      end

    Periods = Args.periods
    AccessModes = Args.AccessModes
    
    for period in Periods do
        for acceMode in AccessModes do
            ok = RunMacro("transit assign", Args, period, acceMode)
            if !ok then goto quit
          end
      end

  quit:
    return(ok)
endmacro

Macro "transit assign" (Args, period, acceMode)
    rsFile = Args.Routes
    tnwFile = Args.(period + " " + acceMode + " Transit Net") // AM_walk.tnw, AM_pnr.tnw, AM_knr.tnw
    odMatrix = Args.(period + " " + acceMode + " OD Matrix")  // AM_walk_PA.mtx, AM_pnr_PA.mtx, AM_knr_PA.mtx

    if acceMode = "walk" then 
        TransModes = Args.TransModes // {"bus", "brt", "rail", "all"}
    else // if "pnr" or "knr"
        TransModes = Subarray(Args.TransModes, 1, 3) // {"bus", "brt", "rail"}

    for transMode in TransModes do
        label = period + " " + acceMode + " " + transMode

        lineFlowTable = Args.(label + " LineFlow Table")
        walkFlowTable = Args.(label + " WalkFlow Table")
        aggrFlowTable = Args.(label + " AggrFlow Table")
        boardVolTable = Args.(label + " Boarding Table")

        ok = RunMacro("Set Transit Network", Args, period, acceMode, transMode)
        if !ok then goto quit

            // do assignment
        className = period + "-" + acceMode + "-" + transMode

        obj = CreateObject("Network.PublicTransportAssignment", {RS: rsFile, NetworkName: tnwFile})

        obj.ODLayerType = "Node"
        obj.Method = "PF"
        obj.AddDemandMatrix({Class: className, Matrix: {MatrixFile: odMatrix, Matrix: "weight", RowIndex: "ProdTAZ", ColumnIndex: "AttrTAZ"}})

        obj.FlowTable             = lineFlowTable
        obj.WalkFlowTable         = walkFlowTable
        obj.TransitLinkFlowsTable = aggrFlowTable
        obj.OnOffTable            = boardVolTable

        ok = obj.Run()
        if !ok then goto quit
      end // for transMode

  quit:
    return(ok)
endmacro