Macro "Reports" (Args)
    RunMacro("Load Link Layer", Args)
    RunMacro("Calculate Daily Fields", Args)
    return(1)
endmacro

/*
This loads the final assignment results onto the link layer.
*/

Macro "Load Link Layer" (Args)

    hwy_dbd = Args.HighwayDatabase
    assn_dir = Args.[Output Folder] + "\\assignment\\roadway\\"
    periods = {"AM", "PM", "OP"}

    {nlyr, llyr} = GetDBLayers(hwy_dbd)

    for period in periods do
        assn_file = assn_dir + "\\" + period + "Flows.bin"

        temp = CreateObject("Table", assn_file)
        field_names = temp.GetFieldNames()
        temp = null
        RunMacro("Join Table To Layer", hwy_dbd, "ID", assn_file, "ID1")
        // {map, {nlyr, llyr}} = RunMacro("Create Map", {file: hwy_dbd})
        tbl = CreateObject("Table", {FileName: hwy_dbd, Layer: llyr})
        for field_name in field_names do
            if field_name = "ID1" then continue
            // Remove the field if it already exists before renaming
            tbl.DropFields(field_name + "_" + period)
            tbl.RenameField({FieldName: field_name, NewName: field_name + "_" + period})
        end

        // Calculate delay by time period and direction
        a_dirs = {"AB", "BA"}
        for dir in a_dirs do

            // Add delay field
            delay_field = dir + "_Delay_" + period
            tbl.AddField({
                FieldName: delay_field,
                Description: "Hours of Delay|(CongTime - FFTime) * Flow / 60"
            })

            // Get data vectors
            v_fft = nz(tbl.ABFreeFlowTime)
            v_ct = nz(tbl.(dir + period + "Time"))
            v_vol = nz(tbl.(dir + "_Flow_" + period))

            // Calculate delay
            v_delay = (v_ct - v_fft) * v_vol / 60
            v_delay = max(v_delay, 0)
            tbl.(delay_field) = v_delay
        end
        tbl = null
    end
endmacro

/*
This macro summarize fields across time period and direction.

The loaded network table will have a volume field for each class that looks like
"AB_Flow_auto_AM". It will also have fields aggregated across classes that look
like "BA_Flow_PM". Direction (AB/BA) and time period (e.g. AM)
will be looped over. Create an array of the rest of the field names to
summarize. e.g. {"Flow_auto", "Flow", "VMT"}.
*/

Macro "Calculate Daily Fields" (Args)

  a_periods = {"AM", "PM", "OP"}
  loaded_dbd = Args.HighwayDatabase
  a_dir = {"AB", "BA"}
  modes = {"drivealone", "carpool", "LTRK", "MTRK", "HTRK"}

  // Add link layer to workspace
  {nlyr, llyr} = GetDBLayers(loaded_dbd)
  llyr = AddLayerToWorkspace(llyr, loaded_dbd, llyr)

  // Calculate non-additive daily fields
  fields_to_add = {
    {"AB_Speed_Daily", "Real", 10, 2,,,, "Slowest speed throughout day"},
    {"BA_Speed_Daily", "Real", 10, 2,,,, "Slowest speed throughout day"},
    {"AB_Time_Daily", "Real", 10, 2,,,, "Highest time throughout day"},
    {"BA_Time_Daily", "Real", 10, 2,,,, "Highest time throughout day"},
    {"AB_VOC_Daily", "Real", 10, 2,,,, "Highest v/c throughout day"},
    {"BA_VOC_Daily", "Real", 10, 2,,,, "Highest v/c throughout day"}
  }
  RunMacro("Add Fields", {view: llyr, a_fields: fields_to_add})
  fields_to_add = null

  for d = 1 to a_dir.length do
    dir = a_dir[d]

    v_min_speed = GetDataVector(llyr + "|", dir + "_Speed_Daily", )
    v_min_speed = if (v_min_speed = null) then 9999 else v_min_speed
    v_max_time = GetDataVector(llyr + "|", dir + "_Time_Daily", )
    v_max_time = if (v_max_time = null) then 0 else v_max_time
    v_max_voc = nz(GetDataVector(llyr + "|", dir + "_VOC_Daily", ))

    for p = 1 to a_periods.length do
      period = a_periods[p]

      v_speed = GetDataVector(llyr + "|", dir + "_Speed_" + period, )
      v_time = GetDataVector(llyr + "|", dir + "_Time_" + period, )
      v_voc = GetDataVector(llyr + "|", dir + "_VOC_" + period, )

      v_min_speed = min(v_min_speed, v_speed)
      v_max_time = max(v_max_time, v_time)
      v_max_voc = max(v_max_voc, v_voc)
    end

    SetDataVector(llyr + "|", dir + "_Speed_Daily", v_min_speed, )
    SetDataVector(llyr + "|", dir + "_Time_Daily", v_max_time, )
    SetDataVector(llyr + "|", dir + "_VOC_Daily", v_max_voc, )
  end

  // Sum up the flow fields
  for mode in modes do

    for dir in a_dir do
      out_field = dir + "_" + mode + "_Flow_Daily"
      fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,,"Daily " + dir + " " + mode + " Flow"}}
      v_output = null

      // For this direction and mode, sum every period
      for period in a_periods do
        input_field = dir + "_Flow_" + mode + "_" + period
        v_add = GetDataVector(llyr + "|", input_field, )
        v_output = nz(v_output) + nz(v_add)
      end

      output.(out_field) = v_output
      output.(dir + "_Flow_Daily") = nz(output.(dir + "_Flow_Daily")) + v_output
      output.Total_Flow_Daily = nz(output.Total_Flow_Daily) + v_output
    end

    output.("Total_" + mode + "_Flow_Daily") = output.("AB_" + mode + "_Flow_Daily") + output.("BA_" + mode + "_Flow_Daily")
  end
  fields_to_add = fields_to_add + {
    {"AB_Flow_Daily", "Real", 10, 2,,,,"AB Daily Flow"},
    {"BA_Flow_Daily", "Real", 10, 2,,,,"BA Daily Flow"},
    {"Total_Flow_Daily", "Real", 10, 2,,,,"Daily Flow in both direction"},
  }
  for mode in modes do
    fields_to_add = fields_to_add + {
        {"Total_" + mode + "_Flow_Daily", "Real", 10, 2,,,,"Daily " + mode + " Flow in both direction"},
    }
  end

  // Other fields to sum
  a_fields = {"VMT", "VHT", "Delay"}
  for field in a_fields do
    for dir in a_dir do
      v_output = null
      out_field = dir + "_" + field + "_Daily"
      fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,,"Daily " + dir + " " + field}}
      for period in a_periods do
        input_field = dir + "_" + field + "_" + period
        v_add = GetDataVector(llyr + "|", input_field, )
        v_output = nz(v_output) + nz(v_add)
      end
      output.(out_field) = v_output
      output.("Total_" + field + "_Daily") = nz(output.("Total_" + field + "_Daily")) + v_output
    end

	description = "Daily " + field + " in both directions"
	if field = "Delay" then description = description + " (hours)"
    fields_to_add = fields_to_add + {{"Total_" + field + "_Daily", "Real", 10, 2,,,, description}}
  end

  // The assignment files don't have total delay by period. Create those.
  for period in a_periods do
    out_field = "Tot_Delay_" + period
    fields_to_add = fields_to_add + {{out_field, "Real", 10, 2,,,, period + " Total Delay"}}
    {v_ab, v_ba} = GetDataVectors(llyr + "|", {"AB_Delay_" + period, "BA_Delay_" + period}, )
    v_output = nz(v_ab) + nz(v_ba)
    output.(out_field) = v_output
  end

  RunMacro("Add Fields", {view: llyr, a_fields: fields_to_add})
  SetDataVectors(llyr + "|", output, )
  DropLayerFromWorkspace(llyr)
EndMacro