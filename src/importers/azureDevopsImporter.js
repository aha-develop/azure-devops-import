import React from "react";
import azureDevopsClient, {
  getOrganizationInfo,
  getWorkItems,
  getProjectInfo,
} from "./azureDevops";

const importer = aha.getImporter(
  "aha-develop.azure-devops-import.azure-devops"
);

async function authenticate() {
  const authData = await aha.auth("ado", { useCachedRetry: true });
  azureDevopsClient.setToken(authData.token);
}

importer.on({ action: "listFilters" }, async ({}, { identifier, settings }) => {
  return {
    organization: {
      title: "Organization",
      required: true,
      type: "autocomplete",
    },
    project: {
      title: "Project",
      required: true,
      type: "autocomplete"
    }
  };
});

importer.on(
  { action: "filterValues" },
  async ({ filterName, filters }, { identifier, settings }) => {
    await authenticate();

    const { organization } = filters

    switch (filterName) {
      case "organization":
        return await getOrganizationInfo();
      case "project":
        return await getProjectInfo(organization);
    }

    return []
  }
);

importer.on(
  { action: "listCandidates" },
  async ({ filters, nextPage }, { identifier, settings }) => {
    await authenticate();
    const { workItemList, nextPageOffset } = await getWorkItems(
      filters.organization,
      filters.project,
      nextPage || 0
    );

    if (workItemList == "nothing") {
      return {
        records: [],
        nextPage: nextPageOffset,
      };
    }

    if (workItemList) {
      const records = workItemList.map((workItem) => ({
        uniqueId: workItem.id,
        name: workItem.fields["System.Title"],
        type: workItem.fields["System.WorkItemType"],
        state: workItem.fields["System.State"],
        url: workItem._links.html.href,
        description: workItem.fields["System.Description"],
        tags: (workItem.fields["System.Tags"] || "").split("; ")
      }));

      return {
        records: records,
        nextPage: nextPageOffset,
      };
    } else {
      alert("The organization that you enter doesn't exist.");
      return {
        records: [],
        nextPage: null,
      };
    }
  }
);

importer.on(
  { action: "renderRecord" },
  ({ record, onUnmounted }, { identifier, settings }) => {

    return (
      <div style={{ display: "flex", flexDirection: "row", gap: "4px" }}>
        <div style={{ flexGrow: 1 }}>
          <div className="card__row">
            <div className="card__section">
              <div className="card__field">
                <span className="text-muted">
                  {record.type.toUpperCase()} {record.uniqueId}
                </span>
              </div>
            </div>
            <div className="card__section">
              <div className="card__field">
                <a href={aha.sanitizeUrl(record.url)} target="_blank" rel="noopener noreferrer">
                  <i className="text-muted fa-solid fa-external-link"></i>
                </a>
              </div>
            </div>
          </div>
          <div className="card__row">
            <div className="card__section">
              <div className="card__field">
                <a href={aha.sanitizeUrl(record.url)} target="_blank" rel="noopener noreferrer">
                  {record.name}
                </a>
              </div>
            </div>
          </div>
          <div className="card__row">
            <div className="card__section">
              <div className="card__field">
                <aha-pill color="var(--theme-button-pill)">
                  { record.state }
                </aha-pill>
              </div>
            </div>
            <div className="card__section">
              <div className="card__field">
                <div className="card__tags" data-multiple="true" data-readonly="true">
                  { record.tags.map(tag => (
                    tag && <li style={{ backgroundColor: "var(--aha-purple-300)" }}>{tag}</li>
                  )) }
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
);

importer.on(
  { action: "importRecord" },
  async ({ importRecord, ahaRecord }, { identifier, settings }) => {
    ahaRecord.description = `${importRecord.description}<p><a href='${importRecord.url}'>View in Azure DevOps</a></p>`;
    ahaRecord.tagList = importRecord.tags.join(",");

    await ahaRecord.save();
  }
);
