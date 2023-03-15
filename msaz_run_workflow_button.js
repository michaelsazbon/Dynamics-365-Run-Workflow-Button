



async function executeWorkflowonRecords(entityName, Items) {

    var data = entityName + ";" + Items.length;
    var pageInput = {
        pageType: "webresource",
        webresourceName: "msaz_run_workflow_dialog.html",
        data: data
    };
    var navigationOptions = {
        target: 2,
        width: 600,
        height: 450,
        position: 1,
        title: "Run a Workflow"
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function success(data) {

            //console.log(data);
            if (data && data.returnValue) {

                var workflowId = data.returnValue;
                var notificationObj =
                {
                    type: 2,
                    level: 4, //warning
                    message: "The Workflow is running......",
                    showCloseButton: false
                }

                Xrm.App.addGlobalNotification(notificationObj)
                    .then(
                        async function success(notifid) {

                            var results = [];

                            for (var i = 0; i < Items.length; i++) {

                                //Xrm.App.clearGlobalNotification(notifid);

                                //notificationObj.message = "Execution du Workflow en cours (" + (i + 1) + "/" + Items.length + ") ...",

                                Xrm.Utility.showProgressIndicator("The Workflow is running... (" + (i + 1) + "/" + Items.length + ") ...")

                                //notifid = await Xrm.App.addGlobalNotification(notificationObj);

                                var res = await executeWorkflow(workflowId, Items[i]);
                                //console.log(res);
                                if (!res.succes) {
                                    results.push("record id : " + Items[i] + ", error : " + res.error);
                                }

                            }

                            Xrm.App.clearGlobalNotification(notifid);
                            Xrm.Utility.closeProgressIndicator()
                            

                            if (results.length > 0) {
                                Xrm.Navigation.openErrorDialog({
                                    message: "An error occured on one or more records.",
                                    details: "\n\n" + results.join("\n")
                                }).then(
                                    function (success) {
                                        //console.log(success);
                                    },
                                    function (error) {
                                        console.error(error);
                                    });
                            }

                        },
                        function (error) {
                            console.error(error.message);
                            // handle error here
                        });
            }

        },
        function error() {
            console.error(error.message);
            Xrm.Navigation.openAlertDialog({ text: error.message });
        }
    );



}


async function executeWorkflowonRecord(primaryControl) {

var formContext = primaryControl;
var entityId = formContext.data.entity.getId().replace("{", "").replace("}", "");
var entityName = formContext.data.entity.getEntityName();

var data = entityName + ";1";
var pageInput = {
    pageType: "webresource",
    webresourceName: "msaz_run_workflow_dialog.html",
    data: data
};
var navigationOptions = {
    target: 2,
    width: 600,
    height: 450,
    position: 1,
    title: "Run a Workflow"
};
Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
    function success(data) {
        // Handle dialog closed
        //console.log(data);
        if (data && data.returnValue) {

            var workflowId = data.returnValue;
            var notificationObj =
            {
                type: 2,
                level: 4, //warning
                message: "The Workflow is running......",
                showCloseButton: false
            }

            Xrm.App.addGlobalNotification(notificationObj)
                .then(
                    async function success(notifid) {


                        var res = await executeWorkflow(workflowId, entityId);
                        //console.log(res);

                        Xrm.App.clearGlobalNotification(notifid);

                        if (!res.succes) {
                            Xrm.Navigation.openErrorDialog({
                                message: "An error occured  : " + res.error,
                                details: "\n\n" + res.error
                            }).then(
                                function (success) {
                                    //console.log(success);
                                },
                                function (error) {
                                    console.error(error);
                                });
                        }
                        

                    },
                    function (error) {
                        console.error(error.message);
                        // handle error here
                    });
        }

    },
    function error() {
        console.error(error.message);
        Xrm.Navigation.openAlertDialog({ text: error.message });
    }
);

}


function executeWorkflow(workflowId, recordId) {

try {

    var parameters = {};
    parameters.EntityId = recordId;

    return fetch(Xrm.Page.context.getClientUrl() + "/api/data/v9.1/workflows(" + workflowId.replace("}", "").replace("{", "") + ")/Microsoft.Dynamics.CRM.ExecuteWorkflow", {
        method: "POST",
        headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json"
        },
        body: JSON.stringify(parameters)
    }).then(
        function success(response) {
            if (response.status != 204) {
                return response.json().then((json) => { if (response.ok) { return json; } else { throw json.error; } });
            } else {
                return null;
            }

        }
    ).then(function (response) {

        //console.log(response);
        return {
            succes: true
        }

    }).catch(function (error) {
        console.error(error.message);


        return {
            succes: false,
            error: error.message
        }
    });

} catch (e) {


    console.error(e.message);
    return {
        succes: false,
        error: e.message
    }
}
}


