/*
This is a standalone script that calculates the resilience of a network. The
It is a first draft / proof of concept and is not optimized for speed. It
iterates through each link in the network, disables it, and recalculates the
skim. The difference between the skim with the link disabled and the skim with
the link enabled is the "resilience" of the link. The script writes the results
to the link layer.
*/

Macro "test"
    RunMacro("Resiliency", {
        hwy_dbd: "C:\\projects\\Oahu\\working_files\\resiliency\\base_2022_output_layer\\scenario_links.dbd",
        net_query: "D = 1",
        analyze_query: "resiliency = 1"
    })
    ShowMessage("Resiliency analysis complete.")
endmacro

/*
Inputs

  * hwy_dbd
    * String
    * The link layer to analyze
  * net_query
    * Optional string
    * The query/filter used to create the network. 
    * If null, all links are inclued in the network.
  * analyze_query
    * Optional string
    * Which links included in the network should be analyzed. Use this to
      reduce run times by identifying links that don't need to be processed.
    * If null, all links in the network are analyzed.
*/

Macro "Resiliency" (MacroOpts)

    hwy_dbd = MacroOpts.hwy_dbd
    net_query = MacroOpts.net_query
    analyze_query = MacroOpts.analyze_query

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
        SetName: "links_to_analyze",
        Query: analyze_query
    })
    v_ids = link_tbl.ID

    // Create network
    net = CreateObject("Network.Create")
    net.LayerDB = hwy_dbd
    net.Filter = net_query
    net.LengthField = "Length"
    net.AddLinkField({Name: "Time", Field: {"ABAMTime", "BAAMTime"}, IsTimeField: true})
    {drive, folder, name, ext} = SplitPath(hwy_dbd)
    net_file = folder + "\\" + name + ".net"
    net.OutNetworkName = net_file
    net.Run()

    // Calculate base stats (with all links enabled)
    opts.net_file = net_file
    opts.hwy_dbd = hwy_dbd
    opts.llyr = llyr
    base_stats = RunMacro("Iterate", opts)
    link_tbl.BaseSkimTime = base_stats.Sum
    link_tbl.BaseSkimIJCount = base_stats.Count

    // Iterate through the links to analyze, disabling each one, skimming,
    // and reporting the results
    pbar = CreateObject("G30 Progress Bar", "Analyzing links", true, v_ids.length)
    net_update = CreateObject("Network.Update", {Network: net_file})
    for id in v_ids do
        // Because this can take a long time, skip links that already
        // have a result. This allows the process to be started again and pick up
        // where it left off. To start over completely, manually clear out
        // the result fields in the table.
        rh = LocateRecord(llyr + "|", "ID", {id}, {Exact: "true"})
        if llyr.SkimTime <> null then do
            if pbar.Step() then
                Return()
            continue
        end

        // disable link in network
        net_update.DisableLinks({Type: "BySet", Filter: "ID = " + String(id)})
        net_update.Run()

        // Skim and calculate stats
        stats = RunMacro("Iterate", opts)

        // Write results
        llyr.SkimTime = stats.Sum
        llyr.SkimIJCount = stats.Count

        // Re-enable link before next iteration
        net_update.EnableLinks({Type: "BySet", Filter: "ID = " + String(id)})
        net_update.Run()

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

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
