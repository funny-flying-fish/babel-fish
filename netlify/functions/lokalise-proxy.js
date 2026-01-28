/**
 * Netlify Function: Lokalise API Proxy
 * Proxies requests to Lokalise API to avoid CORS issues in the browser.
 */

const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Api-Token',
    'Access-Control-Max-Age': '86400'
};

exports.handler = async (event, context) => {
    // Handle CORS preflight
    if (event.httpMethod === 'OPTIONS') {
        return {
            statusCode: 200,
            headers: corsHeaders,
            body: ''
        };
    }

    // Extract the Lokalise API path from the request
    // Request comes as: /api/lokalise/projects -> forward to api.lokalise.com/api2/projects
    const path = event.path.replace('/api/lokalise', '/api2');
    const queryString = event.rawQuery ? `?${event.rawQuery}` : '';
    const lokaliseUrl = `https://api.lokalise.com${path}${queryString}`;

    try {
        const headers = {
            'Content-Type': 'application/json'
        };

        // Forward the API token
        const apiToken = event.headers['x-api-token'];
        if (apiToken) {
            headers['X-Api-Token'] = apiToken;
        }

        // Build fetch options
        const fetchOptions = {
            method: event.httpMethod,
            headers
        };

        // Include body for non-GET requests
        if (event.httpMethod !== 'GET' && event.httpMethod !== 'HEAD' && event.body) {
            fetchOptions.body = event.body;
        }

        const response = await fetch(lokaliseUrl, fetchOptions);
        const responseBody = await response.text();

        return {
            statusCode: response.status,
            headers: {
                'Content-Type': 'application/json',
                ...corsHeaders
            },
            body: responseBody
        };
    } catch (error) {
        return {
            statusCode: 500,
            headers: {
                'Content-Type': 'application/json',
                ...corsHeaders
            },
            body: JSON.stringify({ error: { message: error.message } })
        };
    }
};
