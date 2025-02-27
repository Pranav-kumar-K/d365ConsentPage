// JavaScript code for Dynamics 365
async function onLoad(executionContext) {
    try {
        // Get the form context
        var formContext = executionContext.getFormContext();

        // Define the tab and section names
        var tabName = "SUMMARY_TAB";
        var sectionName = "SUMMARY_TAB_section_9";
        var subgridName = "Subgrid_new_1";

        // Access the subgrid
        var subgridControl = formContext.ui.tabs.get(tabName)
            .sections.get(sectionName)
            .controls.get(subgridName);

        // Check if the subgrid exists
        if (subgridControl) {
            // Get the current entity GUID dynamically
            var currentEntityId = formContext.data.entity.getId();

            // Define the custom fetch XML with dynamic GUID
            const fetchXml = `
			<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
				<entity name="activitypointer">
					<attribute name="activitytypecode" />
					<attribute name="subject" />
					<attribute name="statecode" />
					<attribute name="prioritycode" />
					<attribute name="modifiedon" />
					<attribute name="activityid" />
					<attribute name="instancetypecode" />
					<attribute name="community" />
					<order attribute="modifiedon" descending="false" />
					<link-entity name="activityparty" from="activityid" to="activityid" link-type="inner" alias="ad">
						<link-entity name="phonecall" from="activityid" to="activityid" link-type="inner" alias="ae">
							<filter type="and">
								<condition attribute="cr64b_createdfrom" operator="eq" uitype="account" value="${currentEntityId}" />
							</filter>
						</link-entity>
					</link-entity>
				</entity>
			</fetch>`;

				

            // Apply the fetch XML filter to the subgrid
            subgridControl.setFilterXml(fetchXml);

            // Refresh the subgrid to apply the filter
            subgridControl.refresh();

            // Wait for the subgrid data to load
            await waitForSubgridData(subgridControl);

            var grid = subgridControl.getGrid();
            var rows = grid.getRows();

            // Initialize an array to store GUIDs
            var activityGuids = [];

            // Iterate through the rows in the subgrid
            rows.forEach(function (row) {
                var entity = row.getData().getEntity();
                var id = entity.getId(); // Get the GUID
                activityGuids.push(id); // Add it to the array
            });

            // Log the list of GUIDs (optional)
            console.log("Activity GUIDs:", activityGuids);

            // Perform any additional actions with the GUIDs here
        } else {
            console.error("Subgrid not found: " + subgridName);
        }
    } catch (error) {
        console.error("Error in onLoad function: ", error);
    }
}

// Helper function to wait for subgrid data to load
function waitForSubgridData(subgridControl) {
    return new Promise((resolve, reject) => {
        try {
            subgridControl.addOnLoad(() => {
                resolve();
            });
        } catch (error) {
            reject(error);
        }
    });
}
