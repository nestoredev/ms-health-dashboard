require('dotenv').config();
const fs = require('fs');
const axios = require('axios');

// Active issue statuses (filter out resolved/historical)
const ACTIVE_STATUSES = [
    'investigating',
    'serviceInterruption',
    'serviceDegradation',
    'extendedRecovery',
    'falsePositive',
    'investigationSuspended'
];

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
        const headers = { Authorization: `Bearer ${accessToken}` };

        // 2. Fetch Health Overviews (with issues expanded)
        const healthResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews?$expand=issues',
            { headers }
        );

        // 3. For each active issue, fetch the posts (updates)
        const services = [];
        
        for (const item of healthResponse.data.value) {
            const activeIssues = (item.issues || []).filter(issue => 
                ACTIVE_STATUSES.includes(issue.status)
            );

            const issuesWithPosts = [];
            
            for (const issue of activeIssues) {
                let posts = [];
                try {
                    const postsResponse = await axios.get(
                        `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues/${issue.id}/posts`,
                        { headers }
                    );
                    posts = (postsResponse.data.value || []).map(post => ({
                        createdAt: post.createdDateTime,
                        content: post.description?.content || post.body?.content || ''
                    })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
                } catch (e) {
                    // Posts endpoint may not exist for all issues
                    console.log(`No posts for ${issue.id}`);
                }

                issuesWithPosts.push({
                    id: issue.id,
                    title: issue.title,
                    startTime: issue.startDateTime,
                    endTime: issue.endDateTime,
                    lastModified: issue.lastModifiedDateTime,
                    status: issue.status,
                    severity: issue.classification || issue.severity,
                    impactDescription: issue.impactDescription,
                    posts
                });
            }

            services.push({
                service: item.service,
                status: item.status,
                id: item.id,
                issues: issuesWithPosts
            });
        }

        const data = {
            lastUpdated: new Date().toISOString(),
            services
        };

        fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
        console.log(`Health data updated. ${services.reduce((n, s) => n + s.issues.length, 0)} active issues.`);

    } catch (error) {
        console.error("Error fetching health data:", error.response ? error.response.data : error.message);
        process.exit(1);
    }
}

getHealthData();
