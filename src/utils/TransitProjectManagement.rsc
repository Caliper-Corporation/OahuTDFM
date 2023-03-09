// MergeRouteSystem
// TC40 create Route System subset ex

/*
Library of tools to create scenario RTS files by extracting from a master
layer and moving to a scenario network.

General approach and notes:
Every route system (RTS) has a stops dbd file. It ends with "S.dbd".
Opening it gives every piece of information needed to use TransCADs
"Create Route from Table" batch macro.

Select from this layer based on project ID.
Export this selection to a new bin file.
Format the new table to look like what is required by the creation macro
  hover over it and hit F1 for help, which shows the table needed
  Includes creating a new field called Node_ID
Loop over each row to get the nearest node from the new layer.
  Log where stops don't have a scenario node nearby.
Place this value into Node_ID.
Run the create-from-table batch macro.
  Compare distance to previous route to check for large deviations.
*/

Macro "test tpm"

  RunMacro("Close All")

  opts = null
  opts.master_rts = "C:\\projects\\Oahu\\repo_new_model\\master\\networks\\master_routes.rts"
  opts.scen_hwy = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\input\\networks\\scenario_links.dbd"
  opts.proj_list = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\TransitProjectList.csv"
  opts.centroid_qry = "[Zone Centroid] = 'Y'"
  opts.link_qry = "HCMType <> null"
  opts.output_rts_file = "test.rts"
  opts.delete_shape_stops = "true"
  RunMacro("Transit Project Management", opts)

  ShowMessage("Done")
EndMacro

/*
Inputs
  MacroOpts
    Named array containing all function arguments

    master_rts
      String
      Full path to the master RTS file
      The RTS file must have a field called ProjID that contains values matching
      the proj_list CSV file.

    scen_hwy
      String
      Full path to the scenario roadway dbd that will have the routes loaded.

    centroid_qry
      String
      Query that defines centroids in the node layer. Centroids will be
      prevented from having stops tagged to them. Routes will also be prevented
      from traveleing through them.

    link_qry
      Optional string
      A selection set of links that routes can use. By default, all links can
      be used.

    proj_list
      String
      Full path to the CSV file containing the list of routes to include

    delete_shape_stops
      Boolean (default true)
      Whether or not to remove shape stops after transferring routes. Should
      be true when creating scenario networks. Setting it to false is helpful
      when transferring master routes to a new link layer.

    output_rts_file
      Optional String
      The file name desired for the output route system.
      Defaults to "ScenarioRoutes.rts".
      Do not include the full path. The route system will always be created
      in the same folder as the scenario roadway file.

Outputs
  Creates a new RTS file in the same folder as scen_hwy
*/

Macro "Transit Project Management" (MacroOpts)

  // To prevent potential problems with view names, open files, etc.
  // close everything before starting.
  RunMacro("Close All")

  // Argument extraction
  master_rts = MacroOpts.master_rts
  scen_hwy = MacroOpts.scen_hwy
  proj_list = MacroOpts.proj_list
  centroid_qry = MacroOpts.centroid_qry
  link_qry = MacroOpts.link_qry
  output_rts_file = MacroOpts.output_rts_file
  delete_shape_stops = MacroOpts.delete_shape_stops

  // Argument checking
  if master_rts = null then Throw("'master_rts' not provided")
  if scen_hwy = null then Throw("'scen_hwy' not provided")
  if proj_list = null then Throw("'proj_list' not provided")
  if centroid_qry = null then Throw("'centroid_qry' not provided")
  centroid_qry = RunMacro("Normalize Query", centroid_qry)
  link_qry = RunMacro("Normalize Query", link_qry)
  if output_rts_file = null then output_rts_file = "ScenarioRoutes.rts"
  if delete_shape_stops = null then delete_shape_stops = "true"

  // Set the output directory to be the same as the scenario roadway
  a_path = SplitPath(scen_hwy)
  out_dir = a_path[1] + a_path[2]
  out_dir = RunMacro("Normalize Path", out_dir)
  output_rts_file = out_dir + "\\" + output_rts_file

  // The steps below will modify the master route system.
  // Make a temp copy to avoid modifying the actual master.
  {temp_rts, temp_hwy} = RunMacro("Copy RTS Files", {
    from_rts: master_rts,
    to_dir: out_dir,
    include_hwy_files: true
  })
  master_rts = temp_rts

  // Update the values of MacroOpts
  MacroOpts.master_rts = master_rts
  MacroOpts.output_rts_file = output_rts_file
  MacroOpts.centroid_qry = centroid_qry
  MacroOpts.link_qry = link_qry
  MacroOpts.out_dir = out_dir
  MacroOpts.delete_shape_stops = delete_shape_stops

  broken_routes = RunMacro("Migrate Route System", MacroOpts)
  if broken_routes <> null then do
    MacroOpts.broken_routes = broken_routes
    RunMacro("Prepare Broken Routes", MacroOpts)
    RunMacro("Export to GTFS", MacroOpts)
    RunMacro("Import from GTFS", MacroOpts) 
    // RunMacro("Merge Route Systems")
  end
  RunMacro("Update Scenario Attributes", MacroOpts)
  RunMacro("Check Scenario Route System", MacroOpts)
  if delete_shape_stops then RunMacro("Remove Shape Stops", MacroOpts)

  // Remove the temp copies
  RunMacro("Delete RTS Files", temp_rts)
  DeleteDatabase(temp_hwy)
EndMacro

/*

*/

macro "Migrate Route System" (MacroOpts)
	proj_list = MacroOpts.proj_list
  master_rts = MacroOpts.master_rts 
	output_rts_file = MacroOpts.output_rts_file
	scen_hwy = MacroOpts.scen_hwy
	out_dir = MacroOpts.out_dir
		

  // Export the scenario routes into a new RTS
  // TODO: do this with the Table class once the improvements are migrated
  // DataManager.CopyRouteSystem requires a filter (and not a selection set).
  // This means we have to create a 1-hot field to easily select which routes
  // to export.
  tbl = CreateObject("Table", master_rts)
  tbl.AddField({FieldName: "sel_temp"})
  tbl = 0
  dm = CreateObject("DataManager")
  dm.AddDataSource("routes", {
    DataType: "RS",
    FileName: master_rts
  })
  layers = dm.GetRouteLayers("routes")
  rlyr = layers.RouteLayer
  
  // Use the project list to create a selection set of scenario routes
  ptbl = CreateObject("Table", proj_list)
  v_pid = ptbl.ProjID
  if TypeOf(v_pid[1]) <> "string" then v_pid = String(v_pid)
  SetLayer(rlyr)
  for pid in v_pid do
    query = "Select * where ProjID = '" + pid + "'"
    n = SelectByQuery("to_export", "more", query)
    if n = 0 then Throw(
      "TPM: Route with ProjID = '" + pid + "'' is not in the master route system."
    )
  end
  v_one = Vector(n, "Double", {{"Constant", 1}})
  SetDataVector(rlyr + "|to_export", "sel_temp", v_one, )
  dm.CopyRouteSystem("routes", {
    TargetRS: output_rts_file,
    Filter: "sel_temp = 1"
  })
  dm = null

  // Delete any existing error log files before modifying the route system
	{drive, folder, name, ext} = SplitPath(output_rts_file)
	link_err_file = drive + folder + name + "l.err"
	if GetFileInfo(link_err_file) <> null then DeleteFile(link_err_file)
	
  // Point route system to scenario link layer
  {nlyr_s, llyr_s} = GetDBLayers(scen_hwy)  
	ModifyRouteSystem(output_rts_file, {{"Geography", scen_hwy, llyr_s}})
	// TODO: develop a way to do this without a message box popping up.
  map = CreateObject("Map", output_rts_file)

	//if error occurred during migrating, parse the error log and report broken routes
	broken_routes = null 
	if GetFileInfo(link_err_file) <> null then do 

		fp = OpenFile(link_err_file, "r")
		while !FileAtEOF(fp) do 
			line = ReadLine(fp)

			pos1 = Position(line, "route ID")
			pos2 = Position(line, "cannot be found")

			if pos1 > 0 & pos2 > 0 & pos2 > pos1 then do 
				route_id = Trim(SubString(line, pos1 + 9, pos2 - pos1-9))

				if ArrayPosition(broken_routes, {route_id},) <= 0 then
					broken_routes = broken_routes + {s2i(route_id)}
				end 
			end
		CloseFile(fp)
  end 

	return(broken_routes)
endMacro 

/*
Removes broken from the migrated route system and prepares a new transit project
list for just broken routes. This new list will use the gtfs import/export
procedures.
*/

Macro "Prepare Broken Routes" (MacroOpts)
	proj_list = MacroOpts.proj_list
  master_rts = MacroOpts.master_rts 
	output_rts_file = MacroOpts.output_rts_file
	broken_routes = MacroOpts.broken_routes

  // Open the route system and create a selection set of broken routes
  // TODO: make silent
  map = CreateObject("Map", output_rts_file)
  {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
  // TODO: use the table method once tc9 is updated
  SetLayer(rlyr)
  SelectByIDs("broken_routes", "several", broken_routes)
  routes = CreateObject("Table", rlyr)
  routes.ChangeSet("broken_routes")

  // Get the broken route ids and delete them from the route system
  v_proj_ids = routes.ProjID
  v_names = routes.Route_Name
  // if TypeOf(v_proj_ids[1]) <> "string" then v_proj_ids = String(v_proj_ids)
  for name in v_names do
    DeleteRoute(rlyr, name)
  end

  // Create a new projet list that is only of the broken routes
  tbl = CreateObject("Table", proj_list)
  for pid in v_proj_ids do
    tbl.SelectByQuery({
      SetName: "broken_routes",
      Query: "ProjID = " + pid,
      Operation: "more"
    })
  end
  proj_list2 = Substitute(proj_list, ".csv", "_2.csv", )
  tbl.Export({FileName: proj_list2})
  Throw()
endmacro

/*

*/

Macro "Export to GTFS" (MacroOpts)

  master_rts = MacroOpts.master_rts
  proj_list = MacroOpts.proj_list
  scen_hwy = MacroOpts.scen_hwy

  // Get project IDs from the project list
  proj = OpenTable("projects", "CSV", {proj_list, })
  proj_tbl = CreateObject("Table", proj_list)
  v_pid = proj_tbl.ProjID
  if TypeOf(v_pid) = "null" then Throw("No transit project IDs found")
  if TypeOf(v_pid[1]) <> "string" then v_pid = String(v_pid)

  // Convert the project IDs into route IDs
  opts = null
  opts.rts_file = master_rts
  opts.v_pid = v_pid
  {v_rid, } = RunMacro("Convert ProjID to RouteID", opts)

  // Select routes based on route ids
  map = CreateObject("Map", master_rts)
  {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
  SetLayer(rlyr)
  export_set = CreateSet("to export")
  n_prev = 0
  for i = 1 to v_rid.length do
    rid = v_rid[i]
    rid = if TypeOf(rid) = "string" then "'" + rid + "'" else String(rid)
    qry = "Select * where Route_ID = " + rid
    operation = if i = 1 then "several" else "more"
    n = SelectByQuery(export_set, operation, qry)
    if n <= n_prev then Throw(T(
      "Transit Manager: Project %s in the project list was not found in the " +
      "master route system.", 
      {v_pid[i]}
    ))
    n_prev = n
  end

  // Create a gtfs folder to hold the export
  {drive, folder, , } = SplitPath(scen_hwy)
  gtfs_dir = drive + folder + "gtfs"
  RunMacro("Create Directory", gtfs_dir)

  // Export
  gtfs = CreateObject("GTFSExporter", {
    RouteFile: master_rts,
    RouteSet: export_set,
    GTFSFolder: gtfs_dir
  })
  gtfs.Export()

endmacro

/*

*/

Macro "Import from GTFS" (MacroOpts)

  scen_hwy = MacroOpts.scen_hwy
  master_rts = MacroOpts.master_rts
  output_rts_file = MacroOpts.output_rts_file
  link_qry = MacroOpts.link_qry
  delete_shape_stops = MacroOpts.delete_shape_stops
  
  // Create a network of the links to use
  // TODO: the GTFS importer can use different networks for different
  // modes, so we could have a rail network or a brt network
  // that only those modes would use. This may not be necessary given the
  // importer already uses the route alignment, though.
  net_file = RunMacro("Create Simple Roadway Net", {
    hwy_dbd: scen_hwy,
    link_qry: link_qry
  })

  {drive, folder, , } = SplitPath(scen_hwy)
  gtfs_dir = drive + folder + "gtfs"

  gtfs = CreateObject("GTFSImporter", {
    RoadDatabase: scen_hwy,
    GTFSFolder: gtfs_dir,
    RouteFile: output_rts_file,
    NetworkFile: net_file,
    RouteBuffer: 50/5280
  })
  gtfs.ServicesFlag = 0
  gtfs.Import({DropPhysicalStops: true})

  // Remove the temp gtfs directory
  files = RunMacro("Catalog Files", {dir: gtfs_dir})
  for file in files do
    DeleteFile(file)
  end
  RemoveDirectory(gtfs_dir)

  // Create map with both route systems
  map = CreateObject("Map", output_rts_file)
  {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
  {master_rlyr, master_slyr, , } = map.AddRouteLayers({RouteFile: master_rts})

  // Clean up route attributes
  tbl = CreateObject("Table", rlyr)
  tbl.Route_Name = tbl.[Short Name]
  ReloadRouteSystem(output_rts_file)
  tbl.RenameField({FieldName: "Route", NewName: "Master_Route_ID"})
  tbl.ChangeField({
    FieldName: "Master_Route_ID",
    Description: "Route ID from the master route system"
  })
  tbl.DropFields({FieldNames: {
    "Mode",
    "Short Name",
    "Long Name",
    "Description",
    "URL",
    "Color",
    "Text Color",
    "Trip",
    "Sign",
    "Service",
    "Agency Name",
    "Agency URL",
    "Agency Phone",
    "Direction",
    "M", "Tu", "W", "Th", "F", "Sa", "Su",
    "ScheduleStartTime",
    "ScheduleEndTime",
    "AM", "Midday", "PM", "Night",
    "Start Time", "End Time",
    "Headway"
  }})

  // Add back route attributes from the master route system
  fields = {
    {FieldName: "ProjID", Type: "string"},
    {FieldName: "Mode", Type: "integer"},
    {FieldName: "Fare", Type: "real"},
    {FieldName: "EA_Headway", Type: "integer"},
    {FieldName: "AM_Headway", Type: "integer"},
    {FieldName: "MD_Headway", Type: "integer"},
    {FieldName: "PM_Headway", Type: "integer"},
    {FieldName: "EV_Headway", Type: "integer"}
  }
  tbl.AddFields({Fields: fields})
  master_tbl = CreateObject("Table", master_rlyr)
  join_tbl = tbl.Join({
    Table: master_tbl,
    LeftFields: "Master_Route_ID",
    RightFields: "Route_ID"
  })
  join_tbl.(rlyr + ".ProjID") = join_tbl.(master_rlyr + ".ProjID")
  join_tbl.(rlyr + ".Mode") = join_tbl.(master_rlyr + ".Mode")
  join_tbl.(rlyr + ".Fare") = join_tbl.(master_rlyr + ".Fare")
  join_tbl.(rlyr + ".EA_Headway") = join_tbl.(master_rlyr + ".EA_Headway")
  join_tbl.(rlyr + ".AM_Headway") = join_tbl.(master_rlyr + ".AM_Headway")
  join_tbl.(rlyr + ".MD_Headway") = join_tbl.(master_rlyr + ".MD_Headway")
  join_tbl.(rlyr + ".PM_Headway") = join_tbl.(master_rlyr + ".PM_Headway")
  join_tbl.(rlyr + ".EV_Headway") = join_tbl.(master_rlyr + ".EV_Headway")

  // Clean up stop attributes
  stop_tbl = CreateObject("Table", slyr)
  stop_tbl.DropFields({FieldNames: {
    "Pickup",
    "Dropoff",
    "Service",
    "Length",
    "Sequence"
  }})
  // Add back stop attributes from master
  fields = {
    {FieldName: "shape_stop", Type: "integer"}//,
    // TODO: transfer these fields once they exist on the stop layer
    // {FieldName: "dwell_on", Type: "real"},
    // {FieldName: "dwell_off", Type: "real"},
    // {FieldName: "xfer_pen", Type: "real"}
  }
  stop_tbl.AddFields({Fields: fields})
  stop_tbl.RenameField({FieldName: "GTFS_Stop_ID", NewName:"Master_Stop_ID"})
  master_tbl = CreateObject("Table", master_slyr)
  join_tbl = stop_tbl.Join({
    Table: master_tbl,
    LeftFields: "Master_Stop_ID",
    RightFields: "ID"
  })
  join_tbl.(slyr + ".shape_stop") = join_tbl.(master_slyr + ".shape_stop")
  // TODO: transfer these fields once they exist on the stop layer
  // join_tbl.(slyr + ".dwell_on") = join_tbl.(master_slyr + ".dwell_on")
  // join_tbl.(slyr + ".dwell_off") = join_tbl.(master_slyr + ".dwell_off")
  // join_tbl.(slyr + ".xfer_pen") = join_tbl.(master_slyr + ".xfer_pen")
  join_tbl = null
endmacro

/*
Updates the scenario route system attributes based on the TransitProjectList.csv
Also tags stops with node IDs in the 'node_id' field.
*/

Macro "Update Scenario Attributes" (MacroOpts)

  // Argument extraction
  scen_hwy = MacroOpts.scen_hwy
  proj_list = MacroOpts.proj_list
  output_rts_file = MacroOpts.output_rts_file

  // Read in the parameter file
  param = CreateObject("Table", proj_list)

  // Create a map of the scenario RTS
  map = CreateObject("Map", output_rts_file)
  {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()

  SetLayer(rlyr)

  // Loop over column names and update attributes. ProjID is skipped.
  a_field_names = param.GetFieldNames()
  // Only do this process if columns other than ProjID exist.
  if a_field_names.length > 1 then do
    for field_name in a_field_names do
      if field_name = "ProjID" then continue

      // Filter out null values from this column
      n = param.SelectByQuery({
        SetName: "no_nulls",
        Query: field_name + " <> null"
      })
      if n = 0 then continue

      v_pid = param.ProjID
      v_value = param.(field_name)
      for i = 1 to v_pid.length do
        pid = v_pid[i]
        value = v_value[i]

        // Locate the route with this project ID. If not found, throw an error.
        opts = null
        opts.Exact = "true"
        rh = LocateRecord(rlyr + "|", "ProjID", {String(pid)}, opts)
        if rh = null then do
          pid_string = if TypeOf(pid) = "string" then pid else String(pid)
          Throw("ProjID '" + pid_string + "' not found in route layer")
        end

        // Update the attribute
        SetRecord(rlyr, rh)
        rlyr.(field_name) = value
      end
    end
  end

  // Tag stops to nodes within
  a_field = {{"Node_ID", "Integer", 10, , , , , "ID of node closest to stop"}}
  RunMacro("Add Fields", {view: slyr, a_fields: a_field})
  n = TagRouteStopsWithNode(rlyr,,"Node_ID",.2)
EndMacro

/*
Creates a CSV with stats on how well the route system transfer did
*/

Macro "Check Scenario Route System" (MacroOpts)

  // Argument extraction
  master_rts = MacroOpts.master_rts
  scen_hwy = MacroOpts.scen_hwy
  proj_list = MacroOpts.proj_list
  centroid_qry = MacroOpts.centroid_qry
  output_rts_file = MacroOpts.output_rts_file
  out_dir = MacroOpts.out_dir

  // Create path to the copy of the master rts and roadway files
  {drive, path, filename, ext} = SplitPath(master_rts)
  opts = null
  opts.rts_file = master_rts
  master_hwy_copy = RunMacro("Get RTS Roadway File", opts)

  // Get project IDs from the project list and convert to route ids on both
  // the master and scenario route systems.
  proj = OpenTable("projects", "CSV", {proj_list, })
  v_pid = GetDataVector(proj + "|", "ProjID", )
  if TypeOf(v_pid) = "null" then Throw("No transit project IDs found")
  if TypeOf(v_pid[1]) <> "string" then v_pid = String(v_pid)
  CloseView(proj)
  opts = null
  opts.rts_file = master_rts
  opts.v_pid = v_pid
  {v_rid_m, v_rev_pid} = RunMacro("Convert ProjID to RouteID", opts)
  opts.rts_file = output_rts_file
  {v_rid_s, } = RunMacro("Convert ProjID to RouteID", opts)

  // Open the master and scenario route systems in separate maps.
  master_map = CreateObject("Map", master_rts)
  {nlyr_m, llyr_m, rlyr_m, slyr_m} = master_map.GetLayerNames()
  scen_map = CreateObject("Map", output_rts_file)
  {nlyr_s, llyr_s, rlyr_s, slyr_s} = scen_map.GetLayerNames()
  
  // Compare master and scenario routes
  data = null
  for i = 1 to v_rev_pid.length do
    pid = v_rev_pid[i]
    rid_m = v_rid_m[i]
    rid_s = v_rid_s[i]

    // Calculate the route length in the master rts
    opts = null
    opts.rlyr = rlyr_m
    opts.llyr = llyr_m
    opts.lookup_field = "Route_ID"
    opts.id = rid_m
    {length_m, num_stops_m} = RunMacro("Get Route Length and Stops", opts)
    // Calculate the route length in the scenario rts
    opts = null
    opts.rlyr = rlyr_s
    opts.llyr = llyr_s
    opts.lookup_field = "Route_ID"
    opts.id = rid_s
    {length_s, num_stops_s} = RunMacro("Get Route Length and Stops", opts)

    // calculate difference and percent difference
    diff = length_s - length_m
    pct_diff = round(diff / length_m * 100, 2)

    // store this information in a named array
    data.projid = data.projid + {pid}
    data.master_route_id = data.master_route_id + {rid_m}
    data.scenario_route_id = data.scenario_route_id + {rid_s}
    data.master_length = data.master_length + {length_m}
    data.scenario_length = data.scenario_length + {length_s}
    data.length_diff = data.length_diff + {diff}
    data.pct_diff = data.pct_diff + {pct_diff}
    data.master_stops = data.master_stops + {num_stops_m}
    data.scenario_stops = data.scenario_stops + {num_stops_s}
    data.missing_stops = data.missing_stops + {num_stops_m - num_stops_s}
  end

  // Convert the named array into a table
  fields = {{FieldName: "projid", Type: "string"}}
  for i = 2 to data.length do
    field_name = data[i][1]
    fields = fields + {{FieldName: field_name}}
  end
  comp_tbl = CreateObject("Table", {Fields: fields})
  comp_tbl.AddRows({EmptyRows: data.projid.length})
  for i = 1 to data.length do
    field_name = data[i][1]
    values = data.(field_name)
    comp_tbl.(field_name) = values
  end
  comp_tbl.Export({FileName: out_dir + "/_rts_creation_results.csv"})
EndMacro

/*
Helper macro used to convert project IDs (which are on the route layer) to
route IDs (which are included on the node and link tables of the route layer).

Inputs
  MacroOpts
    Named array that holds arguments (e.g. MacroOpts.master_rts)

    rts_file
      String
      Full path to the .rts file that contains both route and project IDs

    v_pid
      Array or vector of unique project IDs.

Returns
  An array of two vectors
  A vector of route IDs corresponding to the input project IDs.
  A revised vector of project IDs. If Project IDs are associated with more
    than one route, the project ID will be repeated in this vector.
*/

Macro "Convert ProjID to RouteID" (MacroOpts)

  // Argument extraction
  rts_file = MacroOpts.rts_file
  v_pid = MacroOpts.v_pid

  // Check that the project IDs are unique
  opts = null
  opts.Unique = "true"
  v_unique = SortVector(v_pid, opts)
  if v_pid.length <> v_unique.length then Throw(
    "'v_pid' must be a vector of unique values"
  )

  // Create map of RTS
  opts = null
  opts.file = rts_file
  {map, {rlyr, slyr, phlyr}} = RunMacro("Create Map", opts)

  // Create an error file that will list any project IDs not found in the
  // route system.
  num_notfound = 0
  {drive, path, , } = SplitPath(rts_file)
  error_file = drive + path + "TransitBuildingError.csv"
  file = OpenFile(error_file, "w")
  WriteLine(file, "Below projects are missing:")

  // Convert project IDs into route IDs
  SetLayer(rlyr)
  route_set = "scenario routes"
  for i = 1 to v_pid.length do
    id = v_pid[i]

    // Select routes that match the current project id
    type = TypeOf(id)
    id2 = if type = "string" then "'" + id + "'" else String(id)
    qry = "Select * where ProjID = " + id2
    operation = if i = 1 then "several" else "more"
    n = SelectByQuery(route_set, operation, qry)
    if n = n_prev then do
      num_notfound = num_notfound + 1
      if type = "string"
        then WriteLine(file, id)
        else WriteLine(file, String(id))
    end

    n_prev = n
  end

  CloseFile(file)
  if num_notfound > 0 then Throw(
    "Projects not found in the master route system. See error log in the scenario folder."
  )
  //remove error log once scenario created successfully
  if num_notfound = 0 and GetFileInfo(error_file) <> null then DeleteFile(error_file)

  // Get final results
  v_rid = GetDataVector(rlyr + "|" + route_set, "Route_ID", )
  v_rev_pid = GetDataVector(rlyr + "|" + route_set, "ProjID", )

  CloseMap(map)
  return({v_rid, v_rev_pid})
EndMacro

/*
Helper function for "Check Scenario Route System".
Determines the length of the links that make up a route and number of stops.
If this ends being used by multiple macros in different scripts, move it
to the ModelUtilities.rsc file.

MacroOpts
  Named array that holds other arguments (e.g. MacroOpts.rlyr)

  rlyr
    String
    Name of route layer

  llyr
    String
    Name of the link layer

  lookup_field
    String
    Field to search for the given ID

  id
    String or Integer
    ID to look for in lookup_field

Returns
  Array of two items
    * Length of route
    * Number of stops on route
*/

Macro "Get Route Length and Stops" (MacroOpts)

  // Argument extraction
  rlyr = MacroOpts.rlyr
  llyr = MacroOpts.llyr
  lookup_field = MacroOpts.lookup_field
  id = MacroOpts.id

  // Determine the current layer before doing work to set it back after the
  // macro finishes.
  cur_layer = GetLayer()

  // Get route name based on the lookup_field and id given
  SetLayer(rlyr)
  opts = null
  opts.Exact = "true"
  rh = LocateRecord(rlyr + "|", lookup_field, {id}, opts)
  if rh = null then Throw("Value not found in" + lookup_field)
  SetRecord(rlyr, rh)
  route_name = rlyr.Route_Name
  a_stops = GetRouteStops(rlyr, route_name, )
  num_stops = a_stops.length

  // Get IDs of links that the route runs on
  a_links = GetRouteLinks(rlyr, route_name)
  for link in a_links do
    a_lid = a_lid + {link[1]}
  end

  // Determine length of those links
  SetLayer(llyr)
  n = SelectByIDs("route_links", "several", a_lid, )
  if n = 0 then Throw("Route links not found in layer '" + llyr + "'")
  v_length = GetDataVector(llyr + "|route_links", "Length", )
  length = VectorStatistic(v_length, "Sum", )

  // Set the layer back to the original if there was one.
  if cur_layer <> null then SetLayer(cur_layer)

  return({length, num_stops})
EndMacro

/*
Removes any shape stops
*/

Macro "Remove Shape Stops" (MacroOpts)

  output_rts_file = MacroOpts.output_rts_file
  
  map = CreateObject("Map", output_rts_file)
  {nlyr, llyr, rlyr, slyr} = map.GetLayerNames()
  stop_tbl = CreateObject("Table", slyr)
  stop_fields = stop_tbl.GetFieldNames()
  if delete_shape_stops and stop_fields.position("shape_stop") <> 0 then do
    n = stop_tbl.SelectByQuery({
      SetName: "to remove",
      Query: "shape_stop = 1",
      Operation: "several"
    })
    if n > 0 then do
      SetLayer(stop_tbl.GetView())
      DeleteRecordsInSet("to remove")
    end
  end
endmacro