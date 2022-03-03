import React from "react";
import azureDevopsClient, { getOrganizationInfo, getWorkItems } from './azureDevops';

const importer = aha.getImporter("aha-develop.azure-devops-import.azure-devops");

async function authenticate() {
  const authData = await aha.auth("ado", { useCachedRetry: true });
  azureDevopsClient.setToken(authData.token);
}

importer.on({ action: "listFilters" }, async ({}, {identifier, settings}) => {
  await authenticate();
  await getOrganizationInfo();
  return {
    organization: {
      title: "Organization",
      required: true,
      type: "select",
    },
  };
});


importer.on({ action: "filterValues" }, async ({ filterName, filters }, {identifier, settings}) => {
  let values = [];
  switch (filterName) {
    case "organization":
      values = await getOrganizationInfo();
  }
  return values;
});

importer.on({ action: "listCandidates" }, async ({ filters, nextPage }, {identifier, settings}) => {
  await authenticate();
  const { workItemList, nextPageOffset } = await getWorkItems(filters.organization, nextPage || 0);
  console.log(workItemList);

  if(workItemList == 'nothing') {
    return {
      records: [],
      nextPage: nextPageOffset
    };
  }

  if(workItemList){
      const records = workItemList.map((workItem) => ({
        uniqueId: workItem.id,
        name: workItem.fields['System.Title'],
        state: workItem.fields['System.State'],
        url: workItem._links.html.href,
        description: workItem.fields['System.Description']
      }));

      return {
        records: records ,
        nextPage: nextPageOffset
      };
  }else{
    alert('The organization that you enter doesn\'t exist.');
    return {
      records: [] ,
      nextPage: null
    };
  }

});

importer.on({ action: "renderRecord" }, ({ record, onUnmounted }, { identifier, settings }) => {
  onUnmounted(() => {
    console.log("Un-mounting component for", record.identifier);
  });

  return (
    <div>
      <h6>{record.state}</h6>
      <a href={`${record.url}`}>{record.name}</a>
    </div>
  );
});

importer.on({ action: "importRecord" }, async ({ importRecord, ahaRecord }, {identifier, settings}) => {
  await authenticate();
  ahaRecord.description =
    `${importRecord.description}<p><a href='${importRecord.url}'>View on Azure Devops</a></p>`;
  await ahaRecord.save();
});

