window.onLoadStageChangeEvent = function (executionContext) {
    console.log("📌 onLoadStageChangeEvent triggered");

    var formCtx = executionContext.getFormContext();

    waitForBpfReady(formCtx, function () {
        monitorStageAndBindNext(formCtx);
    });
};

function waitForBpfReady(formCtx, callback) {
    var retries = 0;
    var maxRetries = 20;

    var interval = setInterval(function () {
        var process = formCtx.data.process;
        if (process && process.getActiveStage && process.getActiveStage()) {
            console.log("✅ BPF ready. Stage:", process.getActiveStage().getName());
            clearInterval(interval);
            callback();
        } else {
            retries++;
            console.log("⏳ Waiting for BPF initialization... Attempt", retries);
            if (retries >= maxRetries) {
                clearInterval(interval);
                console.warn("⚠️ BPF not ready after max retries.");
            }
        }
    }, 500);
}

function monitorStageAndBindNext(formCtx) {
    var processingDeptStageIds = [
        "3b5a344f-9f9d-466b-aa08-611e60964b46",
        "851023ff-3126-48b0-9b8c-8d2da2cfa3c5",
        "26b4c9c0-c630-4540-81d3-bf3090fa3c3a"
    ];

    setInterval(function () {
        try {
            var stage = formCtx.data.process.getActiveStage();
            if (!stage) return;

            var stageId = stage.getId().replace(/[{}]/g, "").toLowerCase();
            var nextBtn = window.top.document.querySelector('button[aria-label="Next Stage"]');
            if (!nextBtn) return;

            if (!processingDeptStageIds.includes(stageId)) {
                nextBtn.dataset.modalAttached = "";
                return;
            }

            if (nextBtn.dataset.modalAttached !== "true") {
                console.log("✅ Attaching modal to Next Stage for 'Processing- Department'");
                nextBtn.dataset.modalAttached = "true";

                nextBtn.addEventListener("click", function (e) {
                    e.preventDefault();
                    e.stopPropagation();

                    if (!window.top.document.getElementById("statusCommentModal")) {
                        openCommentModal(formCtx);
                    }
                }, true);
            }
        } catch (err) {
            console.error("💥 Error in modal attach loop:", err.message);
        }
    }, 1000);
}

function injectBootstrapCss() {
    var head = window.top.document.head;
    if (!head.querySelector("#bootstrap-css")) {
        var link = window.top.document.createElement("link");
        link.id = "bootstrap-css";
        link.rel = "stylesheet";
        link.href = "https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css";
        head.appendChild(link);
    }
}

function openCommentModal(formContext) {
    injectBootstrapCss();

    var parentDoc = window.top.document;
    var existing = parentDoc.getElementById("statusCommentModal");
    if (existing) existing.remove();

    var modalHtml = '' +
        '<div id="statusCommentModal" class="modal fade show" tabindex="-1" style="' +
        'background-color: rgba(0,0,0,0.5); position: fixed; top: 0; left: 0;' +
        'width: 100%; height: 100%; z-index: 1055; display: flex; justify-content: center; align-items: center;">' +
        '<div class="modal-dialog modal-dialog-centered" style="max-width: 500px; width: 100%;">' +
        '<div class="modal-content shadow-lg border-0 rounded-3">' +
        '<div class="modal-header">' +
        '<h5 class="modal-title">Status Change Comment</h5>' +
        '<button type="button" class="btn-close" onclick="window.top.closeStatusCommentModal()"></button>' +
        '</div>' +
        '<div class="modal-body">' +
        '<textarea id="statusCommentText" class="form-control w-100" style="min-height: 120px;" placeholder="Enter your comment..."></textarea>' +
        '</div>' +
        '<div class="modal-footer justify-content-end">' +
        '<button class="btn btn-primary px-4" onclick="window.top.submitStatusComment()">Submit</button>' +
        '</div>' +
        '</div></div></div>';

    var container = parentDoc.createElement("div");
    container.innerHTML = modalHtml;
    parentDoc.body.appendChild(container);

    window.top._statusCommentContext = formContext;

    window.top.submitStatusComment = function () {
        var comment = window.top.document.getElementById("statusCommentText").value.trim();
        if (!comment) {
            alert("Please enter a comment.");
            return;
        }

        var formContext = window.top._statusCommentContext;
        if (!formContext) return;

        formContext.data.save().then(function () {
            console.log("✅ Form saved");

            var caseId = formContext.data.entity.getId().replace(/[{}]/g, "");
            var statusAttr = formContext.getAttribute("statuscode");
            var statusLabel = statusAttr && statusAttr.getText ? statusAttr.getText() : statusAttr.getValue();

            var note = {
                "subject": "Stage Change Comment",
                "notetext": "[" + statusLabel + "] " + comment,
                "objectid_incident@odata.bind": "/incidents(" + caseId + ")"
            };

            Xrm.WebApi.createRecord("annotation", note).then(function (result) {
                console.log("📝 Note created:", result.id);

                setCaseOwnerToCustomerService(formContext);

                setTimeout(function () {
                    console.log("⏳ Attempting stage change after save delay...");
                    activateProcessingStage(formContext);
                }, 1000);

                window.top.closeStatusCommentModal();
            }, function (error) {
                console.error("❌ Failed to save note:", error.message);
                alert("Error saving comment: " + error.message);
            });
        }).catch(function (error) {
            console.error("❌ Error saving form before stage change:", error.message);
            alert("Error saving form: " + error.message);
        });
    };

    window.top.closeStatusCommentModal = function () {
        var existing = window.top.document.getElementById("statusCommentModal");
        if (existing) existing.remove();
        window.top._statusCommentContext = null;
    };
}

// ✅ Fallback Owner Assignment (Dev/Prod)
function setCaseOwnerToCustomerService(formContext) {
    const caseId = formContext.data.entity.getId();
    if (!caseId) return;

    const teamGuids = [
        "fca3c311-074c-f011-a400-fbb6a348b744",  // Prod
        "2c80efda-7c4b-f011-a3ff-af212fee8ea9"
    ];

    function tryAssignToTeam(index) {
        if (index >= teamGuids.length) {
            console.error("❌ All team assignment attempts failed.");
            return;
        }

        const teamId = teamGuids[index];
        console.log(`🔄 Trying to assign case to Team [${index + 1}] with GUID: ${teamId}`);

        Xrm.WebApi.updateRecord("incident", caseId, {
            "ownerid@odata.bind": `/teams(${teamId})`
        }).then(function () {
            console.log(`✅ Case assigned to Team [${index + 1}]`);
            formContext.data.refresh(false);
        }).catch(function (error) {
            console.warn(`⚠️ Failed to assign to Team [${index + 1}]: ${error.message}`);
            tryAssignToTeam(index + 1); // Fallback to next team
        });
    }

    tryAssignToTeam(0); // Start with first GUID
}

function activateProcessingStage(formContext) {
    const targetStageName = "Processing";

    try {
        const activePath = formContext.data.process.getActivePath();
        let targetStage = null;

        for (let i = 0; i < activePath.length; i++) {
            const s = activePath[i];
            if (s.getName().trim().toLowerCase() === targetStageName.toLowerCase()) {
                targetStage = s;
                break;
            }
        }

        if (!targetStage) {
            console.warn("⚠ Stage not found in UI path. Using Web API.");
            forceChangeViaWebAPI(formContext, targetStageName);
            return;
        }

        formContext.data.process.setActiveStage(targetStage.getId(), function (result) {
            if (result === "success") {
                console.log("✅ Stage changed to:", targetStageName);
            } else {
                forceChangeViaWebAPI(formContext, targetStageName);
            }
        });
    } catch (err) {
        console.error("❌ Error in activateProcessingStage:", err.message);
    }
}

function forceChangeViaWebAPI(formContext, targetStageName) {
    try {
        const instanceId = formContext.data.process.getInstanceId();
        const processId = formContext.data.process.getActiveProcess().getId();
        if (!instanceId || !processId) return;

        Xrm.WebApi.retrieveRecord("workflow", processId, "?$select=uniquename").then(function (workflow) {
            const bpfEntityLogicalName = workflow.uniquename.toLowerCase();

            Xrm.WebApi.retrieveRecord(bpfEntityLogicalName, instanceId, "?$expand=processid($select=workflowid)").then(function (bpfRecord) {
                const actualProcessId = bpfRecord.processid.workflowid;

                Xrm.WebApi.retrieveMultipleRecords("processstage", `?$filter=processid/workflowid eq ${actualProcessId}`).then(function (stageResults) {
                    let matchedStage = null;
                    for (let i = 0; i < stageResults.entities.length; i++) {
                        const stage = stageResults.entities[i];
                        if (stage.stagename.trim().toLowerCase() === targetStageName.toLowerCase()) {
                            matchedStage = stage;
                            break;
                        }
                    }

                    if (!matchedStage) return;

                    const updateData = {
                        "activestageid@odata.bind": `/processstages(${matchedStage.processstageid})`
                    };

                    Xrm.WebApi.updateRecord(bpfEntityLogicalName, instanceId, updateData).then(function () {
                        console.log("✅ Stage updated via Web API.");
                    }, function (err) {
                        console.error("❌ Web API stage update failed:", err.message);
                    });
                });
            });
        });
    } catch (err) {
        console.error("❌ forceChangeViaWebAPI failed:", err.message);
    }
}
