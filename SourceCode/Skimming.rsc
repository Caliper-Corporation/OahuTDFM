macro "HighwayAndTransitSkim" (Args, Result)
    ret_value = 1

    ret_value = RunMacro("HighwayNetworkSkim", Args)
    if !ret_value then goto quit
    ret_value = RunMacro("TransitNetworkSkim", Args)
    if !ret_value then goto quit

    quit:
    Return(ret_value)

endmacro

macro "HighwayNetworkSkim" (Args)
    ret_value = 1

    LineDB = Args.HighwayDatabase
    netfile = Args.HighwayNetwork
    AMhwyskimfile = Args.HighwaySkimAM
    PMhwyskimfile = Args.HighwaySkimPM
    OPhwyskimfile = Args.HighwaySkimOP
    walkskimfile = Args.WalkSkim
    bikeskimfile = Args.BikeSkim
    
    Line = CreateObject("Table", {FileName: LineDB, LayerType: "Line"})
    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    NodeLayer = Node.GetView()

    TAZData = Args.DemographicOutputs
    dem = CreateObject("Table", TAZData)

    SkimFiles = {AMhwyskimfile, PMhwyskimfile, OPhwyskimfile}
    SkimVar = {"AMTime", "PMTime", "OPTime"}
    for i = 1 to SkimFiles.length do
        hwyskimfile = SkimFiles[i]
        skimvar = SkimVar[i]
        obj = CreateObject("Network.Skims")
        obj.LoadNetwork (netfile)
        obj.LayerDB = LineDB
        obj.Origins ="Centroid <> null"
        obj.Destinations = "Centroid <> null"
        obj.Minimize = skimvar
        obj.AddSkimField({"Length", "All"})
        obj.OutputMatrix({MatrixFile: hwyskimfile, Matrix: "HighwaySkim"})
        ok = obj.Run()

        obj = null
        obj = CreateObject("Distribution.Intrazonal")
        obj.SetMatrix({MatrixFile:hwyskimfile, Matrix: skimvar})
        obj.OperationType = "Replace"
        obj.TreatMissingAsZero = false
        obj.Neighbours = 3
        obj.Factor = 0.5
        ret_value = obj.Run()
        if !ret_value then goto quit

        obj = null
        obj = CreateObject("Distribution.Intrazonal")
        obj.SetMatrix({MatrixFile:hwyskimfile, Matrix: "Length (Skim)"})
        obj.OperationType = "Replace"
        obj.TreatMissingAsZero = false
        obj.Neighbours = 3
        obj.Factor = 0.5
        ok = obj.Run()

        m = CreateObject("Matrix", hwyskimfile)
        currCoreNames = m.GetCoreNames()
        m.RenameCores({CurrentNames: currCoreNames, NewNames: {"Time", "Distance"}})
        idx = m.AddIndex({IndexName: "TAZ",
                    ViewName: NodeLayer, Dimension: "Both",
                    OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
        idxint = m.AddIndex({IndexName: "InternalTAZ",
                    ViewName: NodeLayer, Dimension: "Both",
                    OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
            
    end


    obj = CreateObject("Network.Skims")
    obj.LoadNetwork (netfile)
    obj.LayerDB = LineDB
    obj.Origins ="Centroid <> null"
    obj.Destinations = "Centroid <> null"
    obj.Minimize = "WalkTime"
    obj.AddSkimField({"Length", "All"})
    obj.OutputMatrix({MatrixFile: walkskimfile, Matrix: "WalkSkim"})
    ok = obj.Run()

    obj = CreateObject("Network.Skims")
    obj.LoadNetwork (netfile)
    obj.LayerDB = LineDB
    obj.Origins ="Centroid <> null"
    obj.Destinations = "Centroid <> null"
    obj.Minimize = "BikeTime"
    obj.AddSkimField({"Length", "All"})
    obj.OutputMatrix({MatrixFile: bikeskimfile, Matrix: "BikeSkim"})
    ok = obj.Run()

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:walkskimfile, Matrix: "WalkTime"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:walkskimfile, Matrix: "Length (Skim)"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:bikeskimfile, Matrix: "BikeTime"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    obj = null
    obj = CreateObject("Distribution.Intrazonal")
    obj.SetMatrix({MatrixFile:bikeskimfile, Matrix: "Length (Skim)"})
    obj.OperationType = "Replace"
    obj.TreatMissingAsZero = false
    obj.Neighbours = 3
    obj.Factor = 0.5
    ret_value = obj.Run()
    if !ret_value then goto quit

    m = CreateObject("Matrix", walkskimfile)
    m.RenameCores({CurrentNames: {"WalkTime", "Length (Skim)"}, NewNames: {"Time", "Distance"}})
    idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
    idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})

    m = CreateObject("Matrix", bikeskimfile)
    m.RenameCores({CurrentNames: {"BikeTime", "Length (Skim)"}, NewNames: {"Time", "Distance"}})
    idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
    idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})

    quit:
    Return(ret_value)

endmacro

macro "TransitNetworkSkim" (Args)
    ret_value = 1 
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    TransitTNW = Args.TransitNetwork

    Node = CreateObject("Table", {FileName: LineDB, LayerType: "Node"})
    NodeLayer = Node.GetView()


    classes = {"WalkAM", "DriveAM", "WalkPM", "DrivePM", "WalkOP", "DriveOP"}

    WalkSkim = Args.TransitWalkSkim
    DriveSkim = Args.TransitDriveSkim
    WalkSkimAM = Args.TransitWalkSkimAM
    DriveSkimAM = Args.TransitDriveSkimAM
    WalkSkimPM = Args.TransitWalkSkimPM
    DriveSkimPM = Args.TransitDriveSkimPM
    WalkSkimOP = Args.TransitWalkSkimOP
    DriveSkimOP = Args.TransitDriveSkimOP


    SkimMatrices = {WalkSkimAM, DriveSkimAM, WalkSkimPM, DriveSkimPM, WalkSkimOP, DriveSkimOP}
    Impedances = {"TransitTimeAM", "TransitTimeAM", "TransitTimePM", "TransitTimePM", "TransitTimeOP", "TransitTimeOP"}
    for i = 1 to classes.length do
        cls = classes[i]
        Impedance = Impedances[i]

        o = CreateObject("Network.SetPublicPathFinder", {RS: RouteSystem, NetworkName: TransitTNW})
        o.UserClasses = classes
        o.CurrentClass = cls
        o.DriveTime = "Time"
        o.CentroidFilter = "Centroid <> null"
        o.LinkImpedance = Impedance
        o.Parameters({
        MaxTripCost: 999,
        MaxTransfers: 3,
        VOT: 12,
        MidBlockOffset: 1,
        InterArrival: 0.5
        })
        o.AccessControl({
        PermitWalkOnly: false,
        StopAccessField: null,
        MaxWalkAccessPaths: 10,
        WalkAccessNodeField: null
        })
        o.Combination({
        CombinationFactor: 0.1,
        Walk: 0,
        Drive: 0,
        ModeField: null,
        WalkField: null
        })
        o.StopTimeFields({
        InitialPenalty: null,
        TransferPenalty: null,
        DwellOn: null,
        DwellOff: null
        })
        o.RouteTimeFields({
        Headway: "PeakHeadway"
        })
        o.TimeGlobals({
        MaxInitialWait: 30,
        MaxTransferWait: 30,
        MinInitialWait: 2,
        MinTransferWait: 2,
        TransferPenalty: 0,
        DwellOn: 0.1,
        DwellOff: 0.1,
        MaxAccessWalk: 45,
        MaxEgressWalk: 45,
        MaxTransferWalk: 15
        })
        o.GlobalWeights({
        InitialWait: 2,
        TransferWait: 2,
        WalkTimeFactor: 2,
        DriveTimeFactor: 1.0
        })
        o.Fare({
        Type: "Flat", // Flat, Zonal, Mixed
        RouteFareField: "Fare",
        RouteXFareField: "Fare"
        })
        o.DriveAccess({
        InUse: {false, true, false, true, false, true},
        MaxDriveTime: 20,
        MaxParkToStopTime: 5,
        ParkingNodes: "ParkAndRideLot = 1"
        })
        ok = o.Run()

        skimmatrix = SkimMatrices[i]

        obj = CreateObject("Network.PublicTransportSkims")
        obj.Method = "PF"
        obj.LayerRS = RouteSystem
        obj.LoadNetwork( TransitTNW )
        obj.OriginFilter = "Centroid <> null"
        obj.DestinationFilter = "Centroid <> null"
        obj.SkimVariables = {"Fare", "Initial Wait Time","Transfer Wait Time", "Transfer Walk Time",
                                        "Access Walk Time", "Egress Walk Time", "Access Drive Time", "Dwelling Time", "Total Time",
                                        "Number of Transfers","In-Vehicle Time", "Drive Distance"}
        obj.OutputMatrix({MatrixFile: skimmatrix, Matrix: cls + "PTSkim"})
        ok = obj.Run()

        m = CreateObject("Matrix", skimmatrix)
        idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
        idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
        end

    quit:
    Return(ret_value)


endmacro