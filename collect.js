require('dotenv').config();
const fs = require('fs');
const axios = require('axios');

// Active issue statuses
const ACTIVE_STATUSES = [
    'investigating',
    'serviceInterruption',
    'serviceDegradation',
    'extendedRecovery'
];

// Resolved statuses for history
const RESOLVED_STATUSES = [
    'serviceRestored',
    'postIncidentReviewPublished',
    'resolved'
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

        const services = [];
        const recentHistory = [];
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        
        for (const item of healthResponse.data.value) {
            const activeIssues = (item.issues || []).filter(issue => 
                ACTIVE_STATUSES.includes(issue.status)
            );

            // Get resolved issues from past 30 days
            const resolvedIssues = (item.issues || []).filter(issue => {
                if (!RESOLVED_STATUSES.includes(issue.status)) return false;
                const endDate = issue.endDateTime ? new Date(issue.endDateTime) : new Date(issue.lastModifiedDateTime);
                return endDate >= thirtyDaysAgo;
            });

            const issuesWithPosts = [];
            
            // Fetch active issues with full details
            for (const issue of activeIssues) {
                const fullIssue = await fetchIssueDetails(issue.id, headers);
                if (fullIssue) issuesWithPosts.push(fullIssue);
            }

            // Fetch resolved issues with full details (for history)
            for (const issue of resolvedIssues) {
                const fullIssue = await fetchIssueDetails(issue.id, headers, true);
                if (fullIssue) {
                    fullIssue.serviceName = item.service;
                    recentHistory.push(fullIssue);
                }
            }

            services.push({
                service: item.service,
                status: item.status,
                id: item.id,
                issues: issuesWithPosts
            });
        }

        // Sort history by end date (most recent first)
        recentHistory.sort((a, b) => {
            const dateA = new Date(a.endTime || a.lastModified);
            const dateB = new Date(b.endTime || b.lastModified);
            return dateB - dateA;
        });

        const data = {
            lastUpdated: new Date().toISOString(),
            services,
            history: recentHistory.slice(0, 50) // Limit to 50 most recent
        };

        fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
        const totalIssues = services.reduce((n, s) => n + s.issues.length, 0);
        console.log(`Health data updated. ${totalIssues} active issues, ${recentHistory.length} resolved in past 30 days.`);

    } catch (error) {
        console.error("Error fetching health data:", error.response ? error.response.data : error.message);
        process.exit(1);
    }
}

async function fetchIssueDetails(issueId, headers, isHistory = false) {
    try {
        const issueResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues/${issueId}`,
            { headers }
        );
        
        const fullIssue = issueResponse.data;
        
        // Extract info from posts
        let scopeOfImpact = '';
        let rootCause = '';
        let userImpact = fullIssue.impactDescription || '';
        
        if (fullIssue.posts && fullIssue.posts.length > 0) {
            // Get from the most recent post
            const latestPost = fullIssue.posts[fullIssue.posts.length - 1];
            const content = latestPost.description?.content || '';
            
            const scopeMatch = content.match(/Scope of impact:\s*([^]*?)(?=Root cause:|Next update by:|Final status:|$)/i);
            if (scopeMatch) scopeOfImpact = scopeMatch[1].trim();
            
            const rootMatch = content.match(/Root cause:\s*([^]*?)(?=Next update by:|Final status:|Scope of impact:|$)/i);
            if (rootMatch) rootCause = rootMatch[1].trim();
            
            const impactMatch = content.match(/User impact:\s*([^]*?)(?=Current status:|More info:|$)/i);
            if (impactMatch && !userImpact) userImpact = impactMatch[1].trim();
        }
        
        const posts = (fullIssue.posts || []).map(post => ({
            createdAt: post.createdDateTime,
            postType: post.postType,
            content: post.description?.content || ''
        })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

        const result = {
            id: fullIssue.id,
            title: fullIssue.title,
            startTime: fullIssue.startDateTime,
            endTime: fullIssue.endDateTime,
            lastModified: fullIssue.lastModifiedDateTime,
            status: fullIssue.status,
            severity: fullIssue.classification,
            userImpact,
            scopeOfImpact,
            rootCause,
            feature: fullIssue.feature,
            isResolved: fullIssue.isResolved,
            posts: isHistory ? posts.slice(0, 5) : posts // Limit history posts
        };
        
        console.log(`Fetched ${issueId}: ${posts.length} updates${isHistory ? ' (history)' : ''}`);
        return result;
    } catch (e) {
        console.error(`Error fetching ${issueId}:`, e.message);
        return null;
    }
}

getHealthData();
