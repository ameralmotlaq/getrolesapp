const fetch = require('node-fetch').default;

// add role names to this object to map them to group ids in your AAD tenant
const roleGroupMappings = {
    'admin': '4f89bcc5-95e1-4ee3-ab23-e1611398970c'    
};

const appRolesMappings = {
    'Contributors': 'ff16315c-5438-44d1-84b0-3dd7bfd506ff',
    'Viewers': '96113c40-6a7a-4999-a048-184a6aee5b37'
};

module.exports = async function (context, req) {
    const user = req.body || {};
    const roles = [];
    
    try
    {
        for (const [role, groupId] of Object.entries(roleGroupMappings)) {
        if (await isUserInGroup(groupId, user.accessToken)) {
            roles.push(role);
            }
        }

        for (const [role, roleId] of Object.entries(appRolesMappings)) {
            if (await isUserInRole(roleId, user.accessToken)) {
                roles.push(role);
            }
            else
            {
                const customRoles = await GetUserRoles(user.accessToken)
                roles.push(customRoles.value);
                roles.push("NoAppRole");
            }
        }
    }
    catch(ex)
    {
        roles.push(ex);
    }
    
    
    context.res.json({
        roles
    });
}

async function isUserInRole(roleId, bearerToken) {
    const url = new URL('https://graph.microsoft.com/v1.0/me/appRoleAssignments');    
    url.searchParams.append('$filter', `resourceId eq '3fab498e-c15c-405a-b0dd-13ed86a51bc6'`);    
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

async function GetUserRoles(bearerToken) {
    const url = new URL('https://graph.microsoft.com/v1.0/me/appRoleAssignments');    
    url.searchParams.append('$filter', `resourceId eq '3fab498e-c15c-405a-b0dd-13ed86a51bc6'`);    
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
    return graphResponse.value;    
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
