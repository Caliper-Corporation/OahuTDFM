/*

*/

Macro "Highway Assignment AM OP PM" (Args)
    periods = {"AM", "PM", "OP"}
    RunMacro("Highway Assignment", Args, periods)
    return(1)
endmacro

/*

*/

Macro "Highway Assignment" (Args, periods)

    hwy_dbd = Args.HighwayDatabase
    feedback_iter = Args.Iteration
    // vot_param_file = Args.[Input Folder] + "/assignment/vot_params.csv"
    assign_iters = Args.AssignIterations
    net_dir = Args.[Output Folder] + "/skims"

    {drive, folder, name, ext} = SplitPath(Args.AMFlows)
    RunMacro("Create Directory", drive + folder)

    for period in periods do
        od_mtx = Args.(period + "_OD")
        net_file = net_dir + "/highwaynet_" + period + ".net"

        o = CreateObject("Network.Assignment")
        o.Network = net_file
        o.LayerDB = hwy_dbd
        o.ResetClasses()
        o.Iterations = assign_iters
        o.Convergence = Args.AssignmentConvergence
        o.Method = "CUE"
        o.Conjugates = 3
        o.DelayFunction = {
            Function: "bpr.vdf",
            Fields: {"FreeFlowTime", "Capacity", "ALPHA", "BETA", "None"}
        }
        o.DemandMatrix({MatrixFile: od_mtx, RowIndex: "Rows", ColIndex: "Columns"})
        o.MSAFeedback({
            Flow: period + "MSAFlow",
            Time: period + "MSATime",
            Iteration: feedback_iter
        })

        o.FlowTable = Args.(period + "Flows")  
        
        // Add classes for each combination of vehicle type and VOT

        // TODO: set VOI for auto classes
        voi = 1
        o.AddClass({
            Demand: "drivealone",
            PCE: 1,
            VOI: voi,
            ExclusionFilter: "HOV = 'HOV'"
        })
        // hov
        o.AddClass({
            Demand: "carpool",
            PCE: 1,
            VOI: voi
        })
        // Light Trucks
        o.AddClass({
            Demand: "LTRK",
            PCE: 1,
            VOI: voi
        })
        // Medium Trucks
        o.AddClass({
            Demand: "MTRK",
            PCE: 1.5,
            VOI: voi
        })
        // Heavy Trucks
        o.AddClass({
            Demand: "HTRK",
            PCE: 2.5,
            VOI: voi
        })

        ret_value = o.Run()
        results = o.GetResults()
        /*
        Use results.data to get flow rmse and other metrics:
        results.data.[Relative Gap]
        results.data.[Maximum Flow Change]
        results.data.[MSA RMSE]
        results.data.[MSA PERCENT RMSE]
        etc.
        */

        // Update link layer
        line = CreateObject("Table", {FileName: hwy_dbd, LayerType: "Line"})
        flow = CreateObject("Table", Args.(period + "Flows"))
        jv = line.Join({Table: flow, LeftFields: "ID", RightFields: "ID1"})
        jv.("AB" + period + "Time") = jv.AB_MSA_Time
        jv.("BA" + period + "Time") = jv.BA_MSA_Time
        // jv.("AB" + period + "SPEED") = jv.("Length") / jv.AB_MSA_Time * 60.0
        // jv.("BA" + period + "SPEED") = jv.("Length") / jv.BA_MSA_Time * 60.0
        jv = null
        flow = null
        line = null
        
        // Update time field on network
        obj = CreateObject("Network.Update")
        obj.LayerDB = hwy_dbd
        obj.Network = net_file
        obj.UpdateLinkField({Name: "Time", Field: {"AB" + period + "Time", "BA" + period + "Time"}})
        obj.Run()
        obj = null
    end

endmacro

macro "FeedbackConvergence" (Args)
    if Args.Iteration = null then Args.Iteration = 1
    if Args.MaxIterations = null then Args.MaxIterations = 1
    if Args.Iteration < Args.MaxIterations then
        retValue = 2
    else
        retValue = 1
    
    Args.Iteration = Args.Iteration + 1
    Return(retValue)
endmacro