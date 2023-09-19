Macro "run"

    // model_skim_dir = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\Output\\skims"
    // output_dir = "C:\\projects\\Oahu\\repo_new_model\\docs\\data\\input\\skims"
    // skim_files = {
    //     "HighwaySkimAM",
    //     "BikeSkim",
    //     "WalkSkim"
    // }

    // for skim_file in skim_files do
    //     in_file = model_skim_dir + "\\" + skim_file + ".mtx"
    //     out_bin = output_dir + "\\" + skim_file + ".bin"
    //     out_csv = output_dir + "\\" + skim_file + ".csv"

    //     mtx = CreateObject("Matrix", in_file)
    //     mtx.ExportToTable({
    //         FileName: out_bin,
    //         OutputMode: "Table"
    //     })

    //     tbl = CreateObject("Table", out_bin)
    //     tbl.RenameField({FieldName: "Origin", NewName: "p_taz"})
    //     tbl.RenameField({FieldName: "Destination", NewName: "a_taz"})
    //     prefix = if Left(skim_file, 1) = "H" then "auto_"
    //         else if Left(skim_file, 1) = "B" then "bike_"
    //         else "walk_"
    //     tbl.RenameField({FieldName: "Time", NewName: prefix + "time"})
    //     tbl.RenameField({FieldName: "Distance", NewName: prefix + "distance"})
    //     tbl.Export({FileName: out_csv})
    // end

    // // transit
    // model_skim_dir = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\Output\\skims\\transit"
    // output_dir = "C:\\projects\\Oahu\\repo_new_model\\docs\\data\\input\\skims"
    // skim_files = {
    //     "AM_w_bus"
    // }

    // for skim_file in skim_files do
    //     in_file = model_skim_dir + "\\" + skim_file + ".mtx"
    //     out_bin = output_dir + "\\" + skim_file + ".bin"
    //     out_csv = output_dir + "\\" + skim_file + ".csv"

    //     mtx = CreateObject("Matrix", in_file)
    //     mtx.ExportToTable({
    //         FileName: out_bin,
    //         OutputMode: "Table"
    //     })

    //     tbl = CreateObject("Table", out_bin)
    //     tbl.RenameField({FieldName: "RCIndex", NewName: "p_taz"})
    //     tbl.RenameField({FieldName: "RCIndex:1", NewName: "a_taz"})
    //     tbl.RenameField({FieldName: "Generalized Cost", NewName: "bus_gen_cost"})
    //     tbl.RenameField({FieldName: "Fare", NewName: "bus_fare"})
    //     tbl.RenameField({FieldName: "In-Vehicle Time", NewName: "bus_ivtt"})
    //     tbl.RenameField({FieldName: "Initial Wait Time", NewName: "bus_init_wait"})
    //     tbl.RenameField({FieldName: "Transfer Wait Time", NewName: "bus_xfer_wait"})
    //     tbl.RenameField({FieldName: "Initial Penalty Time", NewName: "bus_init_pen_time"})
    //     tbl.RenameField({FieldName: "Transfer Penalty Time", NewName: "bus_xfer_pen_time"})
    //     tbl.RenameField({FieldName: "Transfer Walk Time", NewName: "bus_xfer_walk_time"})
    //     tbl.RenameField({FieldName: "Access Walk Time", NewName: "bus_access_walk_time"})
    //     tbl.RenameField({FieldName: "Egress Walk Time", NewName: "bus_egress_walk_time"})
    //     tbl.RenameField({FieldName: "Dwelling Time", NewName: "bus_dwell_time"})
    //     tbl.RenameField({FieldName: "Total Time", NewName: "bus_total_time"})
    //     tbl.RenameField({FieldName: "Number of Transfers", NewName: "bus_num_xfers"})
    //     tbl.RenameField({FieldName: "In-Vehicle Distance", NewName: "bus_iv_distance"})
    //     tbl.RenameField({FieldName: "Length", NewName: "bus_length"})
    //     tbl.RenameField({FieldName: "WalkTime", NewName: "bus_walk_time"})
    //     tbl.DropFields({FieldNames: {
    //         "Access Drive Time",
    //         "Egress Drive Time",
    //         "In-Vehicle Cost",
    //         "Initial Wait Cost",
    //         "Transfer Wait Cost",
    //         "Initial Penalty Cost",
    //         "Transfer Penalty Cost",
    //         "Transfer Walk Cost",
    //         "Access Walk Cost",
    //         "Egress Walk Cost",
    //         "Access Drive Cost",
    //         "Egress Drive Cost",
    //         "Dwelling Cost",
    //         "Access Drive Distance",
    //         "Egress Drive Distance",
    //         "DriveTime"
    //     }})
    //     tbl.Export({FileName: out_csv})
    // end

    // Logsums
    se_file = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\Output\\sedata\\scenario_se.bin"
    model_ls_dir = "C:\\projects\\Oahu\\repo_new_model\\scenarios\\base_2022\\Output\\visitors\\mc\\logsums"
    out_dir = "C:\\projects\\Oahu\\repo_new_model\\docs\\data\\input\\logsums"

    purposes = {"HBRec", "HBO", "HBEat", "HBShop", "HBW", "NHB"}
    segments = {"personal", "business"}
    for segment in segments do
        for purp in purposes do
            mtx_file = model_ls_dir + "\\logsum_" + purp + "_" + segment + "_AM.mtx"
            mtx = CreateObject("Matrix", mtx_file)
            mtx.AddIndex({
                TableName: se_file,
                Dimension: "Both",
                OriginalID: "TAZ",
                NewID: "TAZ",
                IndexName: "TAZ"
            })
            out_file = out_dir + "\\ls_" + purp + "_" + segment + "_AM.omx"
            CopyMatrix(mtx.Root, {"File Name": out_file, OMX: "true"})
        end
    end

    ShowMessage("Done")
endmacro