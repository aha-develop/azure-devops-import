class azureDevopsClient {
  static _token;

  static setToken = (token) => {
    this._token = token;
  };
}

export default azureDevopsClient;

function getHeader() {
  return {
    Authorization: "Bearer " + azureDevopsClient._token,
    Accept: "application/json",
    "Content-Type": "application/json",
  };
}

function assertAuthenticated(response) {
  if (!response.ok && response.status === 401) {
    throw new aha.AuthError(response.statusText, "ado");
  }

  if (response.redirected && response.url.includes("signin")) {
    throw new aha.AuthError(response.statusText, "ado");
  }
}

async function getJSON(endpoint) {
  const response = await fetch(endpoint, {
    method: "GET",
    headers: getHeader(),
  });

  assertAuthenticated(response);

  return await response.json()
}

export async function getUserID() {
  const response = await getJSON("https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=6.0");
  return response.id;
}

export async function getOrganizationInfo() {
  const userId = await getUserID();
  const response = await getJSON(`https://app.vssps.visualstudio.com/_apis/accounts?memberId=${userId}&api-version=5.1`);

  const organizationInfo = response.value.map((organization) => ({
    text: organization.accountName,
    value: organization.accountName,
  }));
  return organizationInfo;
}

export async function getProjectInfo(organization) {
  if (!organization) {
    return [];
  }

  const response = await getJSON(`https://dev.azure.com/${organization}/_apis/projects?api-version=6.0`);

  const projectInfo = response.value.map((project) => ({
    text: project.name,
    value: project.name,
  }));
  return projectInfo;
}

export async function getWorkItems(organization, project, offset) {
  const body = {
    query: `
      Select  [System.Id], [System.Title], [System.State]
      From WorkItems
      Where [State] NOT IN GROUP 'Done'
      And [System.TeamProject] = '${project}'
      Order by [Microsoft.VSTS.Common.Priority] asc, [System.CreatedDate] desc
    `.replace(/\\n/, ' '),
  };

  let queryResponse = await fetch(
    `https://dev.azure.com/${organization}/_apis/wit/wiql?api-version=5.1`,
    {
      method: "POST",
      headers: getHeader(),
      body: JSON.stringify(body),
    }
  );

  assertAuthenticated(queryResponse);

  let json = await queryResponse.json();
  const workItems = json.workItems.slice(offset, offset + 50);

  if (workItems.length === 0) {
    return { workItemList: "nothing", nextPageOffset: null };
  }
  const workitemsIdStr = workItems.map((workitem) => workitem.id).join(",");
  const workItemsResponse = await getJSON(`https://dev.azure.com/${organization}/_apis/wit/workitems?ids=${workitemsIdStr}&$expand=all&api-version=6.0`);
  return { workItemList: workItemsResponse.value, nextPageOffset: offset + workItems.length };
}
