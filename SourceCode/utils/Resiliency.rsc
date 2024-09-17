/*
This is a standalone script that calculates the resilience of a network. The
It is a first draft / proof of concept and is not optimized for speed. It
iterates through each link in the network, disables it, and recalculates the
skim. The difference between the skim with the link disabled and the skim with
the link enabled is the "resilience" of the link. The script writes the results
to the link layer.
*/

Macro "test"
    hwy_dbd = "C:\\projects\\resiliency\\test_links.dbd"
    RunMacro("Resiliency", {hwy_dbd: hwy_dbd})
    ShowMessage("Resiliency analysis complete.")
endmacro

Macro "Resiliency" (MacroOpts)

    hwy_dbd = MacroOpts.hwy_dbd

    map = CreateObject("Map", hwy_dbd)
    {nlyr, llyr} = map.GetLayerNames()
    link_tbl = CreateObject("Table", llyr)
    link_tbl.AddField("Time")
    link_tbl.AddField({FieldName: "BaseSkimTime", Description: "total skim time with all links enabled"})
    link_tbl.AddField({FieldName: "SkimTime", Description: "total skim time with link disabled"})
    link_tbl.AddField({FieldName: "TimeDiff", Description: "SkimTime - BaseSkimTime"})
    link_tbl.AddField({FieldName: "BaseSkimIJCount", Description: "total skim IJ count with all links enabled"})
    link_tbl.AddField({FieldName: "SkimIJCount", Description: "total skim IJ count with link disabled"})
    link_tbl.AddField({FieldName: "Disconnections", Description: "SkimIJCount - BaseSkimIJCount"})
    link_tbl.Time = link_tbl.Length / link_tbl.PostedSpeed * 60
    link_tbl.SelectByQuery({
        SetName: "not CC",
        Query: "HCMType <> 'CC'"
    })
    v_ids = link_tbl.ID

    // Create network
    net = CreateObject("Network.Create")
    net.LayerDB = hwy_dbd
    net.LengthField = "Length"
    net.AddLinkField({Name: "Time", Field: "Time", IstImeField: true})
    {drive, folder, name, ext} = SplitPath(hwy_dbd)
    net_file = folder + "\\" + name + ".net"
    net.OutNetworkName = net_file
    net.Run()

    opts.net_file = net_file
    opts.hwy_dbd = hwy_dbd
    opts.llyr = llyr
    base_stats = RunMacro("Iterate", opts)
    link_tbl.BaseSkimTime = base_stats.Sum
    link_tbl.BaseSkimIJCount = base_stats.Count
    net_update = CreateObject("Network.Update", {Network: net_file})
    out_file = folder + "\\resilience.csv"
    file = OpenFile(out_file, "w")
    
    for id in v_ids do
        // Because this can take a long time, skip links that already
        // have a result. To start over completely, manually clear out
        // the results fields in the table.
        rh = LocateRecord(llyr + "|", "ID", {id}, {Exact: "true"})
        if llyr.SkimTime <> null then continue

        // disable link in network
        net_update.DisableLinks({Type: "BySet", Filter: "ID = " + String(id)})
        net_update.Run()

        stats = RunMacro("Iterate", opts)

        // Write results
        llyr.SkimTime = stats.Sum
        llyr.SkimIJCount = stats.Count

        // Re-enable link before next iteration
        net_update.EnableLinks({Type: "BySet", Filter: "ID = " + String(id)})
        net_update.Run()
    end

    // Calculate the difference fields
    link_tbl.TimeDiff = link_tbl.SkimTime - link_tbl.BaseSkimTime
    link_tbl.Disconnections = link_tbl.SkimIJCount - link_tbl.BaseSkimIJCount
endmacro

Macro "Iterate" (MacroOpts)
    net_file = MacroOpts.net_file
    hwy_dbd = MacroOpts.hwy_dbd
    llyr = MacroOpts.llyr

    // Skim
    skim = CreateObject("Network.Skims")
    skim.Network = net_file
    skim.LayerDB = hwy_dbd
    skim.Origins = "Centroid = 1"
    skim.Destinations = "Centroid = 1"
    skim.Minimize = "Time"
    mtx_file = GetTempFileName(".mtx")
    skim.OutputMatrix({
        MatrixFile: mtx_file, 
        Matrix: "skim"
    })
    skim.Run()

    // Calculate resilience metrics
    mtx = CreateObject("Matrix", mtx_file)
    stats = mtx.GetMatrixStatistics("Time")
    skim = null
    mtx = null
    DeleteFile(mtx_file)
    return(stats)
endmacro
