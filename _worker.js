export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Handle CORS preflight (OPTIONS) for /api
    if ((url.pathname === '/api' || url.pathname.startsWith('/api')) && request.method === 'OPTIONS') {
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

    // Check if it is an API request (/api or /api/)
    if (url.pathname === '/api' || url.pathname.startsWith('/api')) {
      const GAS_URL = 'https://script.google.com/macros/s/AKfycbyhhEQ86Huaoqu03VYBm9HXk1rnf480dDkPnY13bpPigeGHQUc30JyRUxqF7x4r8zo9/exec';

      const targetUrl = new URL(GAS_URL);
      url.searchParams.forEach((value, key) => {
        targetUrl.searchParams.set(key, value);
      });

      const options = {
        method: request.method,
        redirect: 'follow',
      };

      if (request.method === 'POST' || request.method === 'PUT') {
        options.headers = {
          'Content-Type': 'text/plain;charset=utf-8'
        };
        options.body = await request.text();
      }

      try {
        const response = await fetch(targetUrl.toString(), options);
        const responseText = await response.text();

        return new Response(responseText, {
          status: 200,
          headers: {
            'Content-Type': 'application/json; charset=utf-8',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0'
          },
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

    // Serve static assets with no-cache headers for HTML files to prevent stale caching
    const response = await env.ASSETS.fetch(request);
    const newHeaders = new Headers(response.headers);
    
    // หากเป็นหน้าหลัก หรือไฟล์ HTML ให้บังคับให้เบราว์เซอร์ตรวจสอบไฟล์ใหม่เสมอ (no-cache)
    if (url.pathname === '/' || url.pathname.endsWith('.html') || url.pathname === '') {
      newHeaders.set('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
    }

    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers: newHeaders,
    });
  },
};
