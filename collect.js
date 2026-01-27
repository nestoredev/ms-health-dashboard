const fs = require('fs');
const axios = require('axios');

async function getHealthData() {
    const tenantId = process.env.MS_TENANT_ID;
    const clientId = process.env.MS_CLIENT_ID;
    const clientSecret = process.env.MS_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
        console.error("Missing environment variables");
        process.exit(1);
    }

    try {
        // 1. Get Access Token
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            new URLSearchParams({
                client_id: clientId,
                scope: 'https://graph.microsoft.com/.default',
                client_secret: clientSecret,
                grant_type: 'client_credentials'
            }).toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        // 2. Fetch Health Overviews (with issues expanded)
        const healthResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews?$expand=issues',
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const data = {
            lastUpdated: new Date().toISOString(),
            services: healthResponse.data.value.map(item => ({
                service: item.service,
                status: item.status,
                id: item.id,
                issues: (item.issues || []).map(issue => ({
                    id: issue.id,
                    title: issue.title,
                    startTime: issue.startDateTime,
                    status: issue.status,
                    severity: issue.severity
                }))
            }))
        };

        fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
        console.log("Health data updated successfully.");

    } catch (error) {
        console.error("Error fetching health data:", error.response ? error.response.data : error.message);
        process.exit(1);
    }
}

getHealthData();
