window.onLoadStageChangeEvent = function (executionContext) {
    const formCtx = executionContext.getFormContext();
    console.log("📌 onLoadStageChangeEvent triggered");

    waitForBpfReady(formCtx, () => {
        bindStageChangeListener(formCtx);
        monitorStageAndAttachModal(formCtx);
    });
};

function waitForBpfReady(formCtx, callback) {
    let retries = 0;
    const maxRetries = 20;
    const interval = setInterval(() => {
        const process = formCtx.data.process;
        if (process?.getActiveStage?.()) {
            clearInterval(interval);
            console.log("✅ BPF ready:", process.getActiveStage().getName());
            callback();
        } else if (++retries >= maxRetries) {
            clearInterval(interval);
            console.warn("⚠️ BPF not ready after retries");
        }
    }, 500);
}

function bindStageChangeListener(formCtx) {
    const stageToStatusMap = {
        "15322a8f-67b8-47fb-8763-13a28686c29d": 100000000,
        "92a6721b-d465-4d36-aef7-e8822d7a5a6a": 100000006,
        "3b5a344f-9f9d-466b-aa08-611e60964b46": 1,
        "65894155-4ed9-449b-ab1d-d4d4fb196e48": 100000001,
        "1ee2e3b4-3e83-4fe3-9b5b-490b6e91e8af": 100000002,
        "91153307-982f-479d-af7f-73048b80e52c": 100000008,
        "ef0a2c39-d6d9-4b29-a39b-53dc539f0982": 100000003
    };

    formCtx.data.process.addOnStageChange(() => {
        const stage = formCtx.data.process.getActiveStage();
        const stageId = stage?.getId()?.replace(/[{}]/g, "").toLowerCase();
        const newStatus = stageToStatusMap[stageId];

        if (!stageId || newStatus === undefined) return;

        const statusAttr = formCtx.getAttribute("statuscode");
        if (statusAttr?.getValue() !== newStatus) {
            statusAttr.setValue(newStatus);
            console.log(`✅ Statuscode updated to: ${newStatus}`);
            //formCtx.data.save();
        }
    });
}

function monitorStageAndAttachModal(formCtx) {
    const modalStageIds = ["3b5a344f-9f9d-466b-aa08-611e60964b46"];

    setInterval(() => {
        const stage = formCtx.data.process.getActiveStage();
        const stageId = stage?.getId()?.replace(/[{}]/g, "").toLowerCase();
        const nextBtn = window.top.document.querySelector('button[aria-label="Next Stage"]');

        if (!nextBtn || !stageId) return;

        if (!modalStageIds.includes(stageId)) {
            nextBtn.dataset.modalAttached = "";
            return;
        }

        if (nextBtn.dataset.modalAttached !== "true") {
            nextBtn.dataset.modalAttached = "true";
            console.log("🔗 Attaching modal to Next Stage button");
            nextBtn.addEventListener("click", (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (!window.top.document.getElementById("statusCommentModal")) {
                    openCommentModal(formCtx);
                }
            }, true);
        }
    }, 1000);
}

function injectBootstrapCss() {
    const head = window.top.document.head;
    if (!head.querySelector("#bootstrap-css")) {
        const link = document.createElement("link");
        link.id = "bootstrap-css";
        link.rel = "stylesheet";
        link.href = "https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css";
        head.appendChild(link);
    }
}

function openCommentModal(formContext) {
    injectBootstrapCss();
    const doc = window.top.document;

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

    const wrapper = doc.createElement("div");
    wrapper.innerHTML = modalHtml;
    doc.body.appendChild(wrapper);
    window.top._statusCommentContext = formContext;

    window.top.submitStatusComment = function () {
        const comment = doc.getElementById("statusCommentText")?.value?.trim();
        if (!comment) return alert("Please enter a comment.");

        formContext.data.save().then(() => {
            const caseId = formContext.data.entity.getId().replace(/[{}]/g, "");
            const statusAttr = formContext.getAttribute("statuscode");
            const statusLabel = statusAttr?.getText?.() || statusAttr?.getValue();

            const note = {
                subject: "Stage Change Comment",
                notetext: `[${statusLabel}] ${comment}`,
                "objectid_incident@odata.bind": `/incidents(${caseId})`
            };

            Xrm.WebApi.createRecord("annotation", note).then(() => {
                assignCaseToCustomerService(formContext);
                setTimeout(() => {
                    activateProcessingStage(formContext);
                }, 1000);
                window.top.closeStatusCommentModal();
            }, err => alert("❌ Failed to save comment: " + err.message));
        });
    };

    window.top.closeStatusCommentModal = function () {
        const modal = doc.getElementById("statusCommentModal");
        if (modal) modal.remove();
        window.top._statusCommentContext = null;
    };
}

function assignCaseToCustomerService(formContext) {
    const caseId = formContext.data.entity.getId();
    const teamGuids = [
        "fca3c311-074c-f011-a400-fbb6a348b744", // Production
        "2c80efda-7c4b-f011-a3ff-af212fee8ea9"  // Development
    ];

    function tryAssign(index) {
        if (index >= teamGuids.length) return;

        Xrm.WebApi.updateRecord("incident", caseId, {
            "ownerid@odata.bind": `/teams(${teamGuids[index]})`
        }).then(() => {
            console.log("👥 Assigned to team:", teamGuids[index]);
            formContext.data.refresh(false);
        }).catch(() => tryAssign(index + 1));
    }

    tryAssign(0);
}

// 🔁 Force Processing stage using exact ID
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
				bindStageChangeListener(formCtx);
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
