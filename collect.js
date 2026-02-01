require('dotenv').config();
const fs = require('fs');
const axios = require('axios');

// Active issue statuses - truly active issues only
// Removed investigationSuspended as those are dormant/paused
const ACTIVE_STATUSES = [
    'investigating',
    'serviceInterruption',
    'serviceDegradation',
    'extendedRecovery'
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

        // 3. For each active issue, fetch full details including posts
        const services = [];
        
        for (const item of healthResponse.data.value) {
            const activeIssues = (item.issues || []).filter(issue => 
                ACTIVE_STATUSES.includes(issue.status)
            );

            const issuesWithPosts = [];
            
            for (const issue of activeIssues) {
                try {
                    // Fetch full issue details including posts
                    const issueResponse = await axios.get(
                        `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues/${issue.id}`,
                        { headers }
                    );
                    
                    const fullIssue = issueResponse.data;
                    
                    // Extract scope of impact from the latest post
                    let scopeOfImpact = '';
                    if (fullIssue.posts && fullIssue.posts.length > 0) {
                        const latestPost = fullIssue.posts[fullIssue.posts.length - 1];
                        const content = latestPost.description?.content || '';
                        const scopeMatch = content.match(/Scope of impact:\s*([^]*?)(?=Root cause:|Next update by:|$)/i);
                        if (scopeMatch) {
                            scopeOfImpact = scopeMatch[1].trim();
                        }
                    }
                    
                    const posts = (fullIssue.posts || []).map(post => ({
                        createdAt: post.createdDateTime,
                        postType: post.postType,
                        content: post.description?.content || ''
                    })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

                    issuesWithPosts.push({
                        id: fullIssue.id,
                        title: fullIssue.title,
                        startTime: fullIssue.startDateTime,
                        endTime: fullIssue.endDateTime,
                        lastModified: fullIssue.lastModifiedDateTime,
                        status: fullIssue.status,
                        severity: fullIssue.classification,
                        impactDescription: fullIssue.impactDescription,
                        scopeOfImpact,
                        feature: fullIssue.feature,
                        posts
                    });
                    
                    console.log(`Fetched ${issue.id}: ${posts.length} updates`);
                } catch (e) {
                    console.error(`Error fetching ${issue.id}:`, e.message);
                    issuesWithPosts.push({
                        id: issue.id,
                        title: issue.title,
                        startTime: issue.startDateTime,
                        status: issue.status,
                        posts: []
                    });
                }
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
        const totalIssues = services.reduce((n, s) => n + s.issues.length, 0);
        const totalPosts = services.reduce((n, s) => n + s.issues.reduce((m, i) => m + i.posts.length, 0), 0);
        console.log(`Health data updated. ${totalIssues} active issues, ${totalPosts} total updates.`);

    } catch (error) {
        console.error("Error fetching health data:", error.response ? error.response.data : error.message);
        process.exit(1);
    }
}

getHealthData();
