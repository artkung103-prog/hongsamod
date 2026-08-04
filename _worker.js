export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Handle CORS preflight (OPTIONS) for /api
    if (url.pathname === '/api' && request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type',
          'Access-Control-Max-Age': '86400',
        },
      });
    }

    // Check if it is an API request
    if (url.pathname === '/api') {
      const GAS_URL = 'https://script.google.com/macros/s/AKfycbxbhikalZXE6zvqHzKUNY_RikNsCBkDo0aV5Y0XmmnoGswVMpWh_Dk9xwdMJ6QL-vy3jw/exec';

      const targetUrl = new URL(GAS_URL);
      url.searchParams.forEach((value, key) => {
        targetUrl.searchParams.set(key, value);
      });

      const headers = new Headers();
      // Only copy essential headers to avoid Google security blocking
      const headersToCopy = ['content-type', 'accept', 'accept-language', 'user-agent'];
      headersToCopy.forEach(h => {
        const val = request.headers.get(h);
        if (val) headers.set(h, val);
      });

      const options = {
        method: request.method,
        headers: headers,
        redirect: 'follow',
      };

      if (request.method === 'POST' || request.method === 'PUT') {
        const bodyText = await request.text();
        options.body = bodyText;
      }

      try {
        const response = await fetch(targetUrl.toString(), options);

        // Return response with permissive CORS headers
        const responseHeaders = new Headers();
        responseHeaders.set('Content-Type', response.headers.get('content-type') || 'application/json');
        responseHeaders.set('Access-Control-Allow-Origin', '*');
        responseHeaders.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        responseHeaders.set('Access-Control-Allow-Headers', 'Content-Type');

        const responseBody = await response.text();

        return new Response(responseBody, {
          status: response.status,
          headers: responseHeaders,
        });
      } catch (err) {
        return new Response(JSON.stringify({ status: 'error', message: err.toString() }), {
          status: 500,
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
          },
        });
      }
    }

    // Serve static assets
    return env.ASSETS.fetch(request);
  },
};
