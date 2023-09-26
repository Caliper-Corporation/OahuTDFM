/*

*/

Macro "Highway Assignment AM MD PM" (Args)
    periods = {"AM", "MD", "PM"}
    RunMacro("Highway Assignment", Args, periods)
    return(1)
endmacro

Macro "Highway Assignment NT" (Args)
    periods = {"NT"}
    RunMacro("Highway Assignment", Args, periods)
    return(1)
endmacro

/*

*/

Macro "Highway Assignment" (Args, periods)

    hwy_dbd = Args.HighwayDatabase
    net_file = Args.HighwayNetwork
    feedback_iter = Args.Iteration
    // vot_param_file = Args.[Input Folder] + "/assignment/vot_params.csv"
    assign_iters = Args.AssignIterations

    for period in periods do
        od_mtx = Args.(period + "_OD")

        o = CreateObject("Network.Assignment")
        o.Network = net_file
        o.LayerDB = hwy_dbd
        o.ResetClasses()
        o.Iterations = assign_iters
        o.Convergence = Args.AssignConvergence
        o.Method = "CUE"
        o.Conjugates = 3
        o.DelayFunction = {
            Function: "bpr.vdf",
            Fields: {"FreeFlowTime", period + "Capacity", "ALPHA", "BETA", "None"}
        }
        o.DemandMatrix({MatrixFile: od_mtx, RowIndex: "NodeID", ColIndex: "NodeID"})
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
            VOI: voi
        })
        // hov2
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
        if period = "AM" or period = "MD" or period = "PM" then do
            line = CreateObject("Table", {FileName: hwy_dbd, LayerType: "Line"})
            flow = CreateObject("Table", Args.(period + "Flows"))
            jv = line.Join({Table: flow, LeftFields: "ID", RightFields: "ID1"})
            if period = "AM" or period = "PM" then do
                jv.("AB" + period + "TIME") = jv.AB_MSA_Time
                jv.("BA" + period + "TIME") = jv.BA_MSA_Time
                jv.("AB" + period + "SPEED") = jv.("Length") / jv.AB_MSA_Time * 60.0
                jv.("BA" + period + "SPEED") = jv.("Length") / jv.BA_MSA_Time * 60.0
            end
            else do
                jv.ABOPTIME = jv.AB_MSA_Time
                jv.BAOPTIME = jv.BA_MSA_Time
                jv.ABOPSPEED = jv.("Length") / jv.AB_MSA_Time * 60.0
                jv.BAOPSPEED = jv.("Length") / jv.BA_MSA_Time * 60.0

            end
            jv = null
            flow = null
            line = null
        end
    end
    if periods[1] = "AM" then do
        obj = CreateObject("Network.Update")
        obj.LayerDB = hwy_dbd
        obj.Network = net_file
        obj.UpdateLinkField({Name: "AMTIME", Field: {"ABAMTIME", "BAAMTIME"}})
        obj.UpdateLinkField({Name: "PMTIME", Field: {"ABPMTIME", "BAPMTIME"}})
        obj.UpdateLinkField({Name: "OPTIME", Field: {"ABOPTIME", "BAOPTIME"}})
        obj.Run()
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