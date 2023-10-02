Macro "Reports" (Args)
    RunMacro("Load Link Layer", Args)
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