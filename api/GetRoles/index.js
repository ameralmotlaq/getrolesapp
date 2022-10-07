const fetch = require('node-fetch').default;

// add role names to this object to map them to group ids in your AAD tenant
const roleGroupMappings = {
    'admin': '21a96550-aa02-486e-9297-e6e51b6398fc',
    'reader': '33bb071c-118d-40d1-a5d7-7ced5900b973'
};

const appRolesMappings = {
    'Contributors': 'ff16315c-5438-44d1-84b0-3dd7bfd506ff',
    'Viewers': '96113c40-6a7a-4999-a048-184a6aee5b37'
};

module.exports = async function (context, req) {
    const user = req.body || {};
    const roles = [];
    
    for (const [role, roleId] of Object.entries(appRolesMappings)) {
        if (await isUserInRole(roleId, user.accessToken)) {
            roles.push(role);
        }
    }

    context.res.json({
        roles
    });
}

async function isUserInRole(roleId, bearerToken) {
    const url = new URL('https://graph.microsoft.com/v1.0/me/appRoleAssignments');
    url.searchParams.append('$filter', `appRoleId eq '${roleId}'`);
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${bearerToken}`
        },
    });

    if (response.status !== 200) {
        return false;
    }

    const graphResponse = await response.json();
    const matchingGroups = graphResponse.value.filter(roleAssignment => roleAssignment.appRoleId === roleId);
    return matchingGroups.length > 0;
}

async function isUserInGroup(groupId, bearerToken) {
    const url = new URL('https://graph.microsoft.com/v1.0/me/memberOf');
    url.searchParams.append('$filter', `id eq '${groupId}'`);
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${bearerToken}`
        },
    });

    if (response.status !== 200) {
        return false;
    }

    const graphResponse = await response.json();
    const matchingGroups = graphResponse.value.filter(group => group.id === groupId);
    return matchingGroups.length > 0;
}
