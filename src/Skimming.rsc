/*

*/

macro "HighwayAndTransitSkim Oahu" (Args, Result)
    RunMacro("HighwayNetworkSkim Oahu", Args)
    // RunMacro("TransitNetworkSkim Oahu", Args)
    return(1)
endmacro

macro "HighwayNetworkSkim Oahu" (Args)
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
                    // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                    OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
            
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
                // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})

    m = CreateObject("Matrix", bikeskimfile)
    m.RenameCores({CurrentNames: {"BikeTime", "Length (Skim)"}, NewNames: {"Time", "Distance"}})
    idx = m.AddIndex({IndexName: "TAZ",
                ViewName: NodeLayer, Dimension: "Both",
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})
    idxint = m.AddIndex({IndexName: "InternalTAZ",
                ViewName: NodeLayer, Dimension: "Both",
                // OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null and CentroidType = 'Internal'"})
                OriginalID: "ID", NewID: "Centroid", Filter: "Centroid <> null"})

    quit:
    Return(ret_value)

endmacro