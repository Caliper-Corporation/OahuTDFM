
Macro "Transit Assignment" (Args)
    RunMacro("GenerateTransitOD", Args)
    RunMacro("PTAssign", Args)
    return(1)
endmacro

/*

*/

macro "GenerateTransitOD" (Args)
    ret_value = 1
    periods = {"AM", "PM", "OP"}
    cores = {"AMWalkTransitTrips", "AMDriveTransitTrips", "PMWalkTransitTrips", "PMDriveTransitTrips", "OPWalkTransitTrips", "OPDriveTransitTrips", "DAYWalkTransitTrips", "DAYDriveTransitTrips", "DAYAllTransitTrips"}
    transitod = Args.Transit_OD
    for i = 1 to periods.length do
        per = periods[i]
        odmtx = Args.(per + "_OD")
        mODT = CreateObject("Matrix", odmtx)
        mODT.SetRowIndex("Rows")
        mODT.SetColIndex("Columns")
        if i = 1 then do
           o = CreateObject("Matrix")
            mOut = o.CloneMatrixStructure({MatrixLabel: "TransitTrips", CloneSource: mODT.ptwalk, MatrixFile: transitod, Matrices: cores })
            mo = CreateObject("Matrix", mOut)
            mo.DAYWalkTransitTrips := 0
            mo.DAYDriveTransitTrips := 0
            mo.DAYAllTransitTrips := 0
        end
        mo.(per + "WalkTransitTrips") := mODT.ptwalk
        mo.(per + "DriveTransitTrips") := mODT.ptdrive
        mo.DAYWalkTransitTrips := mo.DAYWalkTransitTrips + nz(mODT.ptwalk)
        mo.DAYDriveTransitTrips := mo.DAYDriveTransitTrips + nz(mODT.ptdrive)
        mo.DAYAllTransitTrips := mo.DAYAllTransitTrips + nz(mODT.ptdrive) + nz(mODT.ptwalk)
    end
    quit:
    Return(ret_value)
endmacro

/*

*/

macro "PTAssign" (Args)
    ret_value = 1
    LineDB = Args.HighwayDatabase
    RouteSystem = Args.TransitRoutes
    net_dir = Args.[Output Folder] + "/skims/transit"
    assn_dir = Args.[Output Folder] + "/Assignment/Transit"
    RunMacro("Create Directory", assn_dir)
    access_modes = Args.AccessModes
    periods = {"AM", "PM", "OP"}
    modeTable = Args.TransitModeTable

    // Get the main transit modes from the mode table. Exclude the all-transit
    // network from assignment.
    transit_modes = RunMacro("Get Transit Net Def Col Names", modeTable)
    transit_modes = ExcludeArrayElements(transit_modes, transit_modes.position("all"), 1)

    // TODO: remove this line to assign pnr after creating separate matrix cores
    access_modes = {"w", "knr"}

    for period in periods do
        for access in access_modes do
            tnet_file = net_dir + "/" + period + "_" + access + ".tnw"
            for transit_mode in transit_modes do
                obj = CreateObject("Network.PublicTransportAssignment", {RS: RouteSystem, NetworkName: tnet_file})
                obj.ODLayerType = "Node"
                obj.Method = "PFE"
                obj.Iterations = 1
                obj.FlowTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_flows.bin"
                obj.WalkFlowTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_walkflows.bin"
                obj.OnOffTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_onoff.bin"
                obj.TransitLinkFlowsTable = assn_dir + "/" + access + "_" + transit_mode + "_" + period + "_agg.bin"
                class_name = period + "-" + access + "-" + transit_mode
                // TODO: update this once OD cores are updated
                access2 = if access = "w"
                    then "Walk"
                    else "Drive"
                mopts = {MatrixFile: Args.Transit_OD, Matrix: period + access2 + "TransitTrips"}
                obj.AddDemandMatrix({Class: class_name, Matrix: mopts})
                ok = obj.Run()
                results = obj.GetResults()
            end
        end
    end

    quit:
    Return(ret_value)
endmacro

